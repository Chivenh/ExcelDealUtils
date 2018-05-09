package com.flytiger.Excel.v316;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import javax.servlet.http.HttpServletResponse;


/**
 * 批量导出EXCEL文件.
 * 
 * @see 使用方式:
 * @see 1.创建对象
 * @see #createBatchWork()
 * @see #createBatchWork(String)
 * @see 2.添加子文件
 * @see #addFile(ExportExcelUtil, String)
 * @see 3.导出压缩文件
 * @see #outZip(HttpServletResponse, String)
 * @author LFH
 * @version 1.0.0
 * @date 2018年3月20日
 */
public final class ExportBatchExcelUtil {
	
	private static final String defaultPath = prop("zipFile.default", "D:/batchFiles/zipFiles/");// 默认压缩文件源目录;
	private static final String folderPath = prop("zipFile.folder", "/batchFiles/zipFiles/");// 替补文件目录.
	private static final String childFolderPath = prop("zipFile.childFolder", "zipFiles");// 子文件夹名.
	private String dirPath;// 当前工作计划压缩文件源目录
	private File dir;// 源目录文件夹
	private int length;// 此压缩文件子文件个数
	private File zip;// 要导出的压缩文件对象.
	private String timeStamp;// 当前时间毫秒值.
	private List<File> exps = new ArrayList<>();// 当前压缩文件中子文件集合.

	@Override
	public String toString() {
		// TODO Auto-generated method stub
		return "压缩文件目录:" + this.dirPath;
	}
	/**
	 * 初始化压缩文件导出的工作对象.
	 * 
	 * @author LFH
	 * @date 2018年3月20日 下午2:29:20
	 * @param dirPath
	 * @return
	 */
	public static ExportBatchExcelUtil createBatchWork(String dirPath) {
		if (!dirPath.endsWith("/")) {
			dirPath += "/";
		}
		if (dirPath.indexOf(childFolderPath) < 0) {
			dirPath = dirPath + folderPath;
		}
		File DIR = new File(dirPath);
		if (!DIR.exists()) {
			DIR.mkdirs();
		}
		return new ExportBatchExcelUtil(DIR);
	}

	/**
	 * 初始化压缩文件导出的工作对象.
	 * 
	 * @author LFH
	 * @date 2018年3月20日 下午2:29:52
	 * @return
	 */
	public static ExportBatchExcelUtil createBatchWork() {
		return createBatchWork(defaultPath);
	}

	/**
	 * 向压缩文件添加子文件
	 * 
	 * @author LFH
	 * @date 2018年3月21日 上午10:01:03
	 * @param file
	 */
	public void addFile(File file) {
		FileOutputStream out = null;
		FileInputStream fis = null;
		String fileName = file.getName();
		int tIndex = fileName.lastIndexOf(".");
		tIndex = tIndex > 0 ? tIndex : fileName.length() - 1;
		String fileType = fileName.substring(tIndex, fileName.length());
		fileName = fileName.substring(0, tIndex);
		fileName = (fileName == null || fileName.trim().length() <= 0 ? "_" : fileName)
				+ this.exps.size() + fileType;// 类似复制文件到可以压缩的目录.
		try {
			fis = new FileInputStream(file);
			out = new FileOutputStream(fileName);
			// 自定义缓冲区对象
			byte[] buf = new byte[1024];
			int by = 0;
			while ((by = fis.read(buf)) != -1) {
				out.write(buf, 0, by);
			}
			this.exps.add(new File(fileName));
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		} finally {
			try {
				if (out != null) {
					out.flush();
					out.close();
				}
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	/**
	 * 向压缩文件添加子Excel文件.
	 * 
	 * @author LFH
	 * @date 2018年3月20日 下午2:30:20
	 * @param exp {@link ExportExcelUtil}(工作表操作对象)
	 * @param fileName {@link String}(文件名不带后缀)
	 */
	public void addFile(ExportExcelUtil exp, String fileName) {
		FileOutputStream out = null;
		fileName = (fileName == null || fileName.trim().length() <= 0 ? "_" : fileName) + this.exps.size() + ".xls";
		try {
			// 以压缩目录为根目录,创建此次要压缩的目录,并创建待压缩的文件输出流.
			out = new FileOutputStream(this.dirPath + fileName);
			// 将工作表写入文件流.
			exp.gettSheet().getWorkbook().write(out);
			// 在待压缩集合中添加此文件对象.
			this.exps.add(new File(this.dirPath, fileName));
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		} finally {
			try {
				if (out != null) {
					out.flush();
					out.close();
				}
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	/**
	 * 批量添加子Excel文件
	 * 
	 * @author LFH
	 * @date 2018年3月21日 上午8:49:02
	 * @param exps {@link Map}<{@link String}(文件名不带后缀),{@link ExportExcelUtil}(工作表操作对象)>
	 */
	public void addFile(Map<String, ExportExcelUtil> exps) {
		for (String fileName : exps.keySet()) {
			this.addFile(exps.get(fileName), fileName);
		}
	}

	/**
	 * 导出压缩文件夹
	 * 
	 * @author LFH
	 * @date 2018年3月20日 下午2:30:59
	 * @param response
	 * @param fileName 文件名不带后缀
	 * @return
	 */
	public int outZip(HttpServletResponse response, String fileName) {
		int t = 0;
		try {
			t = writeZip(response, fileName);
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		return t;
	}

	/**
	 * 对象构造器
	 * 
	 * @author LFH
	 * @date 2018年3月20日
	 * @param dir
	 */
	private ExportBatchExcelUtil(File dir) {
		this.dir = dir;
		this.timeStamp = Thread.currentThread().getId() + "" + Calendar.getInstance().getTimeInMillis();
		// 保存根目录,最后加'/';
		this.dirPath = dir.getAbsolutePath() + "/" + this.timeStamp + "/";
		new File(this.dirPath).mkdir();// 创建压缩用中间文件夹
		this.zip = new File(this.dir, this.timeStamp + ".zip");
	}

	/**
	 * 从本地向响应流写压缩文件.
	 * 
	 * @author LFH
	 * @date 2018年3月20日 下午2:31:20
	 * @param response
	 * @param fileName
	 * @return
	 * @throws Exception
	 */
	private int writeZip(HttpServletResponse response, String fileName) throws Exception {
		ZipFiles();
		OutputStream out = null;
		out = response.getOutputStream();
		response.reset();
		response.setContentType("application/octet-stream;charset=utf-8");
		fileName = (fileName == null || fileName.trim().length() <= 0 ? "_" : fileName);
		fileName = new String(fileName.getBytes("gbk"), "iso-8859-1");
		// 靠这一行向外导出文件,("文件名以双引号包住,是为了避免在火狐等浏览器导出文件时文件名中有空格出现问题的情况 !")
		response.setHeader("Content-disposition", "attachment; filename=\"" + fileName + ".zip" + "\"");
		FileInputStream inStream = new FileInputStream(this.zip);// 获取文件输入流.
		byte[] buf = new byte[4096];// 创建缓存空间.
		int readLength;
		while (((readLength = inStream.read(buf)) != -1)) {
			out.write(buf, 0, readLength);// 写入响应输出流
		}
		inStream.close();// 关闭输入流
		this.deleteFiles();// 删除中间文件.
		if (out != null) {
			out.close();
		}
		return this.length;
	}

	/**
	 * 执行文件压缩.
	 * 
	 * @author LFH
	 * @date 2018年3月20日 下午2:31:40
	 */
	private void ZipFiles() {
		byte[] buf = new byte[1024];
		try {
			ZipOutputStream out = new ZipOutputStream(new FileOutputStream(this.zip));
			this.length = this.exps.size();
			for (int i = 0; i < this.length; i++) {
				// 循环将每个文件压缩写入压缩文件中.
				File fi = this.exps.get(i);
				FileInputStream in = new FileInputStream(fi);
				out.putNextEntry(new ZipEntry(fi.getName()));
				int len;
				while ((len = in.read(buf)) > 0) {
					out.write(buf, 0, len);
				}
				out.closeEntry();
				in.close();
			}
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 在文件被导出后,删除本地压缩文件源文件.
	 * 
	 * @author LFH
	 * @date 2018年3月20日 下午2:31:50
	 */
	private void deleteFiles() {
		File file = this.dir;
		if (file.exists()) {
			File[] files = file.listFiles();
			for (File fi : files) {
				if (fi.exists() && fi.getName().contains(this.timeStamp)) {
					// 针对中间目录,应先清空文件夹,再删除文件夹
					if (fi.isDirectory()) {
						for (File ifi : fi.listFiles()) {
							ifi.delete();
						}
					}
					fi.delete();
				}
			}
		}
	}

	/**
	 * 获取相关系统属性
	 * 
	 * @author LFH
	 * @date 2018年3月20日 下午7:03:11
	 * @return
	 */
	private static String prop(String key, String defaultValue) {
		return System.getProperty(key, defaultValue).trim().replaceAll(" +", "");
	}

}
