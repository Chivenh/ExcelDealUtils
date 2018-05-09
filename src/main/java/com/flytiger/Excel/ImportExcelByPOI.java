package com.flytiger.Excel;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

/**
 * @author LFH
 * @version 1.0.0
 * @date 2017年6月23日
 * @see 导入EXCEL表格.转换成LIST数据输出.
 */
public class ImportExcelByPOI {

	private String[] model;// 模板标题数组
	private String[] outModel;
	private HSSFSheet sheet;
	private Map<String, Object> emptyMap = new HashMap<>();

	/**
	 * @author LFH
	 * @date 2017年6月23日 下午6:30:08
	 * @see 创建此对象的方法
	 * @param model 导入表的标题检查模板
	 * @return
	 */
	public static ImportExcelByPOI createImportExcelByPOI(HSSFSheet sheet, String[] model) {
		if (sheet != null && model != null && model.length > 0) {
			return new ImportExcelByPOI(sheet, model);
		} else {
			return null;
		}
	}

	/**
	 * @author LFH
	 * @date 2017年6月23日 下午6:30:11
	 * @see 创建此对象的方法
	 * @param model 导入表的标题检查模板
	 * @param outModel 输出LIST的键值模板(如果为空,默认和导入表的标题模板一致)
	 * @return
	 */
	public static ImportExcelByPOI createImportExcelByPOI(HSSFSheet sheet, String[] model, String[] outModel) {
		if (sheet != null && model.length > 0 && outModel.length > 0 && model.length == outModel.length) {
			return new ImportExcelByPOI(sheet, model, outModel);
		} else {
			System.err.println("标题长度不一致!");
			return null;
		}
	}

	/**
	 * 判断导入的文件是否正常的.xls文件.
	 * 
	 * @author LFH
	 * @date 2017年11月21日 下午7:14:55
	 * @param file
	 * @return
	 */
	public static ExcelResult isXlsFile(CommonsMultipartFile file) {
		String fileName = file.getOriginalFilename();
		Pattern check = Pattern.compile(".*[^\\.](\\.(xls)|(XLS))$");
		Matcher matcher = check.matcher(fileName);
		boolean b = matcher.matches();// 文件类型符合
		ExcelResult result = new ExcelResult();
		try {
			if (b) {
				InputStream checkStream = file.getInputStream();
				if (!checkFileHead(checkStream, "xls")) {
					result.setSTATUS(ExcelResult.ERROR);
					result.setMSG("文档格式不符,请选择*.xls格式的文件!");
				} else {
					b = true;
				}
			} else {
				result.setSTATUS(ExcelResult.ERROR);
				result.setMSG("文件类型不符br请确认导入EXCEL 97-2003 .xls格式文件!");
			}
		} catch (Exception e) {
			// TODO: handle exception
			b = false;
			result.setSTATUS(ExcelResult.ERROR);
			result.setMSG("文件解析错误!");
		}
		result.setSUCCESS(b);
		return result;
	}

	/**
	 * @author LFH
	 * @date 2017年7月3日 下午11:54:30
	 * @see 检查文档格式是否为.xls格式.
	 * @param fileName 文件名字符串
	 * @return
	 */
	public static boolean checkFileType(String fileName) {
		Pattern check = Pattern.compile(".*[^\\.](\\.(xls)|(XLS))$");
		Matcher matcher = check.matcher(fileName);
		return matcher.matches();
	}

	/**
	 * @author LFH
	 * @date 2017年7月4日 下午2:30:38
	 * @see 根据需要文件类型,判断文件头是否对应.
	 * @param inputStream
	 * @param type
	 * @return
	 */
	public static boolean checkFileHead(InputStream inputStream, String type) {
		byte[] b = new byte[10];
		boolean bb = false;
		try {
			inputStream.read(b, 0, b.length);
			String fileCode = bytesToHexString(b);
			String xlsCode = fileType.getUseFulHex(fileType.value(type));
			boolean bs = (fileCode.toLowerCase().startsWith(xlsCode)) || (xlsCode.toLowerCase().startsWith(fileCode));
			bb = bs;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return bb;
	}

	/**
	 * @author LFH
	 * @date 2017年7月4日 下午2:26:43
	 * @see 文件头转为16进制可比较字符串.
	 * @param src
	 * @return
	 */
	private static String bytesToHexString(byte[] src) {
		StringBuilder stringBuilder = new StringBuilder();
		if (src == null || src.length <= 0) {
			return null;
		}
		for (int i = 0; i < src.length; i++) {
			int v = src[i] & 0xFF;
			String hv = Integer.toHexString(v);
			if (hv.length() < 2) {
				stringBuilder.append(0);
			}
			stringBuilder.append(hv);
		}
		return stringBuilder.toString();
	}

	/**
	 * @author LFH
	 * @date 2017年6月23日 下午7:01:03
	 * @see 检查标题行是否和模板相同.
	 * @param row
	 * @param sc
	 * @return
	 */
	public boolean checkInDom(int row, int sc) {
		HSSFRow hssfRow = this.sheet.getRow(row);
		String[] inDom = theDom(hssfRow, sc);
		return checkInDom(inDom);
	}

	/**
	 * @author LFH
	 * @date 2017年6月23日 下午7:05:26
	 * @see 返回结果LIST
	 * @param rowS 起始行()
	 * @param cellS 起始列
	 * @return
	 */
	public List<Map<String, Object>> transferToList(int rowS, int cellS) {

		return transfer_List(rowS, cellS);
	}

	/**
	 * @author LFH
	 * @date 2017年6月23日
	 * @param model
	 */
	private ImportExcelByPOI(HSSFSheet sheet, String[] model) {
		this.model = model;
		this.outModel = model;
		this.sheet = sheet;
		initEmptyMap();
	}

	private ImportExcelByPOI(HSSFSheet sheet, String[] model, String[] outModel) {
		this.model = model;
		this.outModel = outModel;
		this.sheet = sheet;
		initEmptyMap();
	}

	/**
	 * @author LFH
	 * @date 2017年6月23日 下午11:59:19
	 * @see 初始化空MAP,以作后续判断用.
	 */
	private void initEmptyMap() {
		for (String k : this.outModel) {
			this.emptyMap.put(k, null);
		}
	}

	/**
	 * @author LFH
	 * @date 2017年6月23日 下午6:06:03
	 * @see 检查标题是否对应
	 * @param inDom
	 * @return
	 */
	private boolean checkInDom(String[] inDom) {
		boolean check = true;
		// 比对两个字符串数组,对标题进行比对,确认是否相同,借此判断文档格式;
		if (!Arrays.equals(this.model, inDom)) {
			check = false;
		}
		return check;
	}

	/**
	 * @author LFH
	 * @date 2017年6月23日 下午6:08:02
	 * @see 得到导入表的标题数组
	 * @param row
	 * @param sc
	 * @return
	 */
	private String[] theDom(HSSFRow row, int sc) {
		if (row == null || row.getLastCellNum() < model.length) {
			return null;
		}
		HSSFCell cell = null;
		String[] titles = new String[model.length];
		for (int i = sc; i < sc + model.length; i++) {
			cell = row.getCell(i);
			if (cell != null) {
				titles[i] = cell.getStringCellValue().trim();
			} else {
				titles[i] = null;
			}
		}

		return titles;
	}

	/**
	 * @author LFH
	 * @date 2017年6月23日 下午6:15:15
	 * @see 将某行值组成MAP
	 * @param rowi
	 * @param cellS
	 * @return
	 */
	private Map<String, Object> getCellMap(HSSFRow rowi, int cellS) {
		if (rowi == null) {
			return null;
		}
		if (cellS < 0 || cellS > rowi.getLastCellNum()) {
			return null;
		}
		Map<String, Object> map = new HashMap<String, Object>();
		HSSFCell cell = null;
		Object value = null;
		FormulaEvaluator evaluator = this.sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		for (int i = cellS; i < cellS + this.outModel.length; i++) {
			cell = rowi.getCell(i);
			if (cell == null) {
				value = null;
			} else {
				CellType cellType = cell.getCellTypeEnum();
				if (evaluator != null && cellType == CellType.FORMULA) {
					evaluator.evaluateInCell(cell);
					cellType = cell.getCellTypeEnum();
				}
				if (cellType == CellType.STRING) {
					value = cell.getStringCellValue();
				} else if (cellType == CellType.NUMERIC) {
					// 当单元格格式为数字时,判断是否是日期转化而来,如果是则反转为日期值.
					if (HSSFDateUtil.isCellDateFormatted(cell)) {
						value = cell.getDateCellValue();
					} else {
						value = cell.getNumericCellValue();
					}
				} else if (cellType == CellType.BLANK) {
					value = null;
				} else {
					value = null;
				}
			}
			map.put(this.outModel[i - cellS], value);
		}
		if (map.equals(this.emptyMap)) {
			map = null;
		}
		return map;
	}

	/**
	 * @author LFH
	 * @date 2017年6月23日 下午6:14:55
	 * @see 将表格中值组成LIST
	 * @param sheet
	 * @param rowS 起始行
	 * @param cellS 起始列
	 * @return
	 */
	private List<Map<String, Object>> transfer_List(int rowS, int cellS) {
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		Map<String, Object> map = null;
		HSSFRow row = null;
		int rowE = this.sheet.getLastRowNum();
		;
		if (this.sheet == null || rowS < 0 || rowE < 0) {
			return null;
		}
		for (int i = rowS; i <= rowE; i++) {
			map = new HashMap<String, Object>();
			row = this.sheet.getRow(i);
			map = getCellMap(row, cellS);
			if (map != null && !map.isEmpty()) {
				list.add(map);
			}
		}
		// T_Util.consoleOut("测试List\n "+list.size());
		return list;
	}

	/**
	 * @author LFH
	 * @date 2017年7月4日
	 *       文件类型枚举类
	 */
	private enum fileType {
		XLS("2003EXCEL", "d0cf11e0a1b11ae10000");

		private String typeName;
		private String typeHex;

		private fileType(String typeName, String typeHex) {
			this.typeName = typeName;
			this.typeHex = typeHex;
		}

		public String getTypeHex() {
			return typeHex;
		}

		@SuppressWarnings("unused")
		public String getTypeName() {
			return typeName;
		}

		/**
		 * @author LFH
		 * @date 2017年7月4日 下午2:41:44
		 * @see 根据枚举获取对应HEX码.
		 * @param type
		 * @return
		 */
		public static String getUseFulHex(fileType type) {
			return type.getTypeHex();
		}

		/**
		 * @author LFH
		 * @date 2017年7月4日 下午2:59:55
		 * @see 将字符串转成fileType的类举.
		 * @param type
		 * @return
		 */
		public static fileType value(String type) {
			return fileType.valueOf(type.toUpperCase());
		}
	}

	public static class ExcelResult {
		private String STATUS = "status";
		private String MSG = "msg";
		private boolean SUCCESS = false;
		private static final String ERROR = "error", FAIL = "fail", EMPTY = "empty";

		public String getSTATUS() {
			return STATUS;
		}

		public String getMSG() {
			return MSG;
		}

		public boolean isSUCCESS() {
			return SUCCESS;
		}

		private void setSTATUS(String sTATUS) {
			STATUS = sTATUS;
		}

		private void setMSG(String mSG) {
			MSG = mSG;
		}

		private void setSUCCESS(boolean sUCCESS) {
			SUCCESS = sUCCESS;
		}

		@Override
		public String toString() {
			return "ExcelResult [STATUS=" + STATUS + ", MSG=" + MSG + ", SUCCESS=" + SUCCESS + "]";
		}

		private ExcelResult(String sTATUS, String mSG, boolean sUCCESS) {
			STATUS = sTATUS;
			MSG = mSG;
			SUCCESS = sUCCESS;
		}

		private ExcelResult() {
			// TODO Auto-generated constructor stub
		}

	}
}
