package com.flytiger.Excel;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * 导出CSV格式文件的简单工具类
 * -- 使用入口
 * 
 * @see #create(HttpServletRequest, HttpServletResponse, String, String, String[], List)
 * @see #create(HttpServletRequest, HttpServletResponse, String, String, String[], String[], List)
 * @author LFH
 * @version 1.0.0
 * @date 2018年3月12日
 */
public class ExportCsvUtil {
	private String dirPath;
	private String fileName;
	private String filePath;

	private class CsvData {
		private String[] titles;
		private String[] keys;
		private String data;
	
		private CsvData(String[] titles, String[] keys, List<Map<String, Object>> list) {
			this.titles = titles;
			this.keys = keys;
			getCsvData(list);
		}
	
		private void getCsvData(List<Map<String, Object>> list) {
			List<String> data = new ArrayList<>();
			data.add(String.join(",", this.titles));
			list.forEach(i -> {
				data.add(Arrays.asList(this.keys).stream().map(k -> {
					Object value = i.get(k);
					String v = "";
					if (value == null || "null".equals(value) || "".equals(value)) {
						v = "";
					} else {
						v = value.toString();
					}
					return v;
				}).collect(Collectors.joining(",")));
			});
			this.data = String.join("\r", data);
		}
	}

	/**
	 * 获取web项目绝对路径.
	 * 
	 * @author LFH
	 * @date 2017年11月24日 上午12:56:39
	 * @param request
	 * @return
	 */
	private static String getRootPath(HttpServletRequest request) {
		String rootPath = request.getSession().getServletContext().getRealPath("/");
		return rootPath.replaceAll("\\\\", "/");
	}

	private ExportCsvUtil() {
		// TODO Auto-generated constructor stub
	}

	private ExportCsvUtil(String rootPath, String dirPath, String fileName) {
		this.dirPath = rootPath + dirPath;
		this.fileName = (fileName.indexOf(".csv") > -1 ? fileName : fileName + ".csv");
		this.filePath = rootPath + dirPath + this.fileName;
	}

	private boolean createCsv(CsvData csvData) {
		boolean isSucess = false;
		FileOutputStream out = null;
		OutputStreamWriter osw = null;
		BufferedWriter bw = null;
		File file = new File(this.filePath);
		String dataList = csvData.data;
		try {
			out = new FileOutputStream(file);
			osw = new OutputStreamWriter(out, "gbk");
			bw = new BufferedWriter(osw);
			if (dataList != null && !dataList.isEmpty()) {
				bw.append(new String(new byte[] { (byte) 0xEF, (byte) 0xBB, (byte) 0xBF }));
				bw.append(dataList);
			}
			isSucess = true;
		} catch (Exception e) {
			isSucess = false;
		} finally {
			if (bw != null) {
				try {
					bw.close();
					bw = null;
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (osw != null) {
				try {
					osw.close();
					osw = null;
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (out != null) {
				try {
					out.close();
					out = null;
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return isSucess;
	}

	private CsvData csv(String[] keys, List<Map<String, Object>> list) {
		return new CsvData(keys, keys, list);
	}

	private CsvData csv(String[] titles, String[] keys, List<Map<String, Object>> list) {
		return new CsvData(titles, keys, list);
	}

	private void out(HttpServletResponse response, CsvData csvData) throws Exception {
		File file = null;
		FileInputStream fis = null;
		OutputStream out = null;
		try {
			// 要输出的内容
			response.setContentType("application/csv;charset=UTF-8");
			response.setCharacterEncoding("UTF-8");
			String fname = this.fileName == null || this.fileName.trim().length() <= 0 ? "_" : this.fileName;
			fname = new String(fname.getBytes("gbk"), "iso-8859-1");
			response.setHeader("Content-disposition", "attachment; filename=\"" + fname + "\"");
			if (createCsv(csvData)) {
				file = new File(this.filePath);
				fis = new FileInputStream(file);
				int len = 0;
				byte[] buffer = new byte[1024];
				out = response.getOutputStream();
				out.write(new byte[] { (byte) 0xEF, (byte) 0xBB, (byte) 0xBF });
				while ((len = fis.read(buffer)) > 0) {
					out.write(buffer, 0, len);
				}
				deleteFiles();
			}
		} catch (Exception e) {
			// TODO: handle exception
			throw new Exception("There is some Error,The " + this.fileName + ".csv file is missed!\t");
		} finally {
			if (fis != null) {
				try {
					fis.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			if (out != null) {
				try {
					out.flush();
					out.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}

	}

	private void outStrict(HttpServletResponse response, CsvData csvData) throws Exception {
		FileInputStream fis = null;
		OutputStream out = null;
		OutputStreamWriter osw = null;
		try {
			response.setContentType("application/csv;charset=UTF-8");
			response.setCharacterEncoding("UTF-8");
			String fname = this.fileName == null || this.fileName.trim().length() <= 0 ? "_" : this.fileName;
			fname = new String(fname.getBytes("gbk"), "iso-8859-1");
			response.setHeader("Content-disposition", "attachment; filename=\"" + fname + "\"");
			osw = new OutputStreamWriter(response.getOutputStream(), "gbk");
			// 要输出的内容    
			// osw.write(new String(new byte[] { (byte) 0xEF, (byte) 0xBB, (byte) 0xBF }));
			osw.write(csvData.data);
		} catch (Exception e) {
			// TODO: handle exception
			throw new Exception("There is some Error,The " + this.fileName + ".csv file is missed!\t");
		} finally {
			if (fis != null) {
				try {
					fis.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			if (osw != null) {
				try {
					osw.flush();
					osw.close();
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			if (out != null) {
				try {
					out.flush();
					out.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	
	}

	private static ExportCsvUtil create(HttpServletRequest request, String dirPath,
			String fileName) {
		String rootPath = getRootPath(request);
		return new ExportCsvUtil(rootPath, dirPath, fileName);
	}

	/**
	 * @author LFH
	 * @see 删除该目录filePath下的所有文件(非csv文件)
	 * @param filePath
	 *            文件目录路径
	 */
	private void deleteFiles() {
		File file = new File(this.dirPath);
		if (file.exists()) {
			File[] files = file.listFiles();
			for (int i = 0; i < files.length; i++) {
				if (files[i].isFile() && files[i].getName().contains(this.fileName)) {
					files[i].delete();
				}
			}
		}
	}

	/**
	 * 创建并执行导出csv服务
	 * 
	 * @author LFH
	 * @date 2018年3月12日 下午7:27:42
	 * @param request
	 * @param response
	 * @param dirPath 过程文件目录
	 * @param fileName 导出文件名
	 * @param titles 标题数组
	 * @param keys 键值数组
	 * @param list 数据集合
	 * @throws Exception
	 */
	public static void create(HttpServletRequest request, HttpServletResponse response, String dirPath, String fileName,
			String[] titles, String[] keys, List<Map<String, Object>> list) throws Exception {
		ExportCsvUtil ecs = create(request, dirPath, fileName);
		ecs.outStrict(response, ecs.csv(titles, keys, list));
	}

	/**
	 * 创建并执行导出csv服务
	 * 
	 * @author LFH
	 * @date 2018年3月12日 下午7:27:59
	 * @param request
	 * @param response
	 * @param dirPath 过程文件目录
	 * @param fileName 导出文件名
	 * @param keys 键值数组
	 * @param list 数据集合
	 * @throws Exception
	 */
	public static void create(HttpServletRequest request, HttpServletResponse response, String dirPath, String fileName,
			String[] keys, List<Map<String, Object>> list) throws Exception {
		ExportCsvUtil ecs = create(request, dirPath, fileName);
		ecs.out(response, ecs.csv(keys, list));
	}
}
