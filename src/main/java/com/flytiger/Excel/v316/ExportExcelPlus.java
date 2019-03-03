package com.flytiger.Excel.v316;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.*;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.function.BinaryOperator;

/**
 * ExportExcelPlus
 *
 *  <h2>POI导表工具类(使用POI[3.16+])</h2>
 *  <ul>
 * 		<li>支持 <b>.xls及.xlsx文件格式.</b></li>
 * 		<li> 此类中方法基本已全部修改为对象级方法,即先有初始化,相关方法才能起效</li>
 * 		<li> --添加内容时请同时添加详细注释</li>
 * 		<li> <b>一般使用顺序</b></li>
 * 		<li> {@link #createExportWork(WorkType, HttpServletRequest, String)}</li>
 * 		<li> --初始化</li>
 * 		<li> {@link #entrySheet(List, int, int, String[])} </li>
 * 		<li> --填充值</li>
 * 		<li> {@link #outFile(HttpServletResponse, String)}</li>
 * 		<li> --最后导出文件</li>
 *      <li> <b>更直观使用顺序</b></li>
 * 		<li> {@link WorkType#build(HttpServletRequest, String)}</li>
 * 		<li> --初始化</li>
 * 		<li> {@link #entrySheet(List, int, int, String[])} </li>
 * 		<li> --填充值</li>
 * 		<li> {@link #outFile(HttpServletResponse, String)}</li>
 * 		<li> --最后导出文件</li>
 * 		<li><p>******详细参数请看对应方法注释</p></li>
 * </ul>
 *
 * <p>其它对象方法请参看公共方法列表</p>
 * @version 1.0.0
 * @author LFH
 * @since 2019年03月02日 08:48
 */
public class ExportExcelPlus {

	private Sheet theSheet;
	private Workbook theWorkbook;
	private CellStyle theCellStyle;
	private ConstructorDefine constructorDefine;

	private String rootPath;

	/**
	 * @param isXls 是否xls文档(Excel 2003)
	 * @param tSheet 当前操作工作表
	 * @see org.apache.poi.hssf.usermodel.HSSFSheet
	 * @see org.apache.poi.xssf.usermodel.XSSFSheet
	 * @param tWorkbook 当前文档抽象对象
	 * @see HSSFWorkbook
	 * @see org.apache.poi.xssf.usermodel.XSSFWorkbook
	 * @param rootPath 文档寻址根路径
	 */
	private ExportExcelPlus(boolean isXls, Sheet tSheet, Workbook tWorkbook, String rootPath) {
		this.theSheet = tSheet;
		this.theWorkbook = tWorkbook;
		this.rootPath = rootPath;
		this.theCellStyle = this.theWorkbook.createCellStyle();
		try {
			final Class INT=int.class;
			if(isXls){
				this.constructorDefine = new ConstructorDefine(HSSFRichTextString.class.getConstructor(String.class),
						HSSFClientAnchor.class.getConstructor(INT,INT,INT,INT,short.class,INT,short.class,INT));
			}else{
				this.constructorDefine = new ConstructorDefine(XSSFRichTextString.class.getConstructor(String.class),
						XSSFClientAnchor.class.getConstructor(INT,INT,INT,INT,INT,INT,INT,INT));
			}
		}catch (NoSuchMethodException e){
			e.printStackTrace();
		}

	}

	/**
	 * 初始化sheet
	 *
	 * @param request 当前请求request
	 * @param path Excel文件路径(注意:此path为相对于项目根目录的路径)<br/>
	 *                <b> path 中请勿拼接项目根目录,此工具中会自动去获取项目根目录</b>
	 *             <h4> 若模板保存在非项目所在目录中,请参看-- {@link #createExportWork(WorkType, String, String, int)} </h5>
	 * @author LFH
	 */
	public static ExportExcelPlus createExportWork(WorkType workType, HttpServletRequest request, String path) {
		return createExportWork( workType,getRootPath(request), path, 0);
	}

	/**
	 * 初始化sheet
	 *
	 * @param request 当前请求request
	 * @param path Excel文件路径(注意:此path为相对于项目根目录的路径)<br/>
	 *                <b> path 中请勿拼接项目根目录,此工具中会自动去获取项目根目录</b>
	 *             <h4> 若模板保存在非项目所在目录中,请参看-- {@link #createExportWork(WorkType, String, String, int)} </h5>
	 * @param at 工作表序号(0开始)
	 * @author LFH
	 * @see #createExportWork(WorkType, String, String, int)
	 */
	public static ExportExcelPlus createExportWork(WorkType workType, HttpServletRequest request, String path, int at) {
		return createExportWork( workType,getRootPath(request), path, at);
	}

	/**
	 * 此初始化方法针对于模板并非在项目目录下的情况<br/>
	 * --[要获取多个sheet时,只要第一次初始化了工作对象,后面可以切换和调用]<br/>
	 * {@link #setTheSheet(int)}
	 * --来切换当前sheet
	 * @param rootPath 模板在电脑硬盘中的目录
	 * @param path 模板路径,此路径是模板在rootPath目录之后的路径
	 * @return ExportExcelPlus
	 * @author LFH
	 * @since 2018年3月31日 下午3:51:45
	 */
	public static ExportExcelPlus createExportWork(WorkType workType,String rootPath, String path) {
		return createExportWork(workType,rootPath,path,0);
	}

	/**
	 * 此初始化方法针对于模板并非在项目目录下的情况<br/>
	 * --[要获取多个sheet时,只要第一次初始化了工作对象,后面可以切换和调用]<br/>
	 * {@link #setTheSheet(int)}
	 * --来切换当前sheet
	 * @param rootPath 模板在电脑硬盘中的目录
	 * @param path 模板路径,此路径是模板在rootPath目录之后的路径
	 * @param at 工作表序号(0开始)
	 * @return ExportExcelPlus
	 * @author LFH
	 * @since 2018年3月31日 下午3:51:45
	 */
	public static ExportExcelPlus createExportWork(WorkType workType,String rootPath, String path, int at) {
		Sheet tSheet ;
		Workbook wb ;
		ExportExcelPlus exp ;
		try(FileInputStream fis=fileInput(rootPath, path)) {
				//获取需要的准确类型
				Class<? extends Workbook> workClass = workType.getWorkClass();
				//获取对应构造器
				Constructor<? extends Workbook> workConstructor = workClass
						.getConstructor(InputStream.class);
				wb = workConstructor.newInstance(fis);
				tSheet = wb.getSheetAt(at);
				exp = new ExportExcelPlus(workClass.equals(WorkType.XLS.getWorkClass()),tSheet, wb, rootPath);
		} catch (Exception e) {
				throw new RuntimeException("文件流未获取,请检查路径!");
		}
		return exp;
	}

	/**
	 * <p>
	 *     <span>切换当前sheet,前提是已经创建了模板对象, 或者已克隆生成新sheet {@link #cloneSheet(int, String)}{@link #cloneSheet(int, int)}</span>
	 * </p>
	 * @param at 下标
	 * @author LFH
	 * @since 2017年5月19日 上午10:57:52
	 */
	public void setTheSheet(int at) {
		if (this.theWorkbook == null || this.theWorkbook.getSheetAt(at) == null) {
			System.err.println("当前sheet尚为空,无法实现切换操作!\n请先初始化sheet,或克隆已存在的sheet");
			return;
		}
		Sheet sheet = this.theWorkbook.getSheetAt(at);
		if (sheet == null) {
			throw new RuntimeException("工作表获取异常,请检查下标!");
		} else {
			this.theSheet = sheet;
		}
	}

	/**
	 * 设置sheet名称
	 * @param at 下标
	 * @param name 名称
	 * @author LFH
	 * @since 2017年5月19日 上午11:18:27
	 */
	public void setSheetName(int at, String name) {
		if (this.theWorkbook == null || this.theWorkbook.getSheetAt(at) == null) {
			System.err.println("当前sheet尚为空,无法实现操作");
			return;
		} else if (name == null || name.trim().length() < 1) {
			return;
		}
		this.theWorkbook.setSheetName(at, name);
	}

	public String getRootPath() {
		return rootPath;
	}

	/**
	 * <b>主要用于当表格复杂时,使用此方法获取当前操作工作表进行填充</b>
	 * @return Sheet
	 * @author LFH
	 * @since 2017年5月12日 下午5:34:47
	 */
	public Sheet getTheSheet() {
		return theSheet;
	}

	/**
	 * 保护指定工作表,如果不指定,默认index为当前sheet.
	 *
	 * @param password 密码
	 * @param index 工作表序号
	 * @return boolean
	 * @author LFH
	 * @since 2017年11月28日 下午9:54:47
	 */
	public boolean protectSheet(String password, int... index) {
		boolean b ;
		try {
			if (index != null && index.length > 0) {
				for (int i : index) {
					Sheet xSheet = this.theWorkbook.getSheetAt(i);
					if (xSheet != null) {
						xSheet.protectSheet(password);
					}
				}
			} else {
				this.theSheet.protectSheet(password);
			}
			b = true;
		} catch (Exception e) {
			throw (e);
		}
		return b;
	}


	/**
	 * <b>克隆一个sheet</b>
	 * @param at 原sheet位置
	 * @return 返回克隆成功的sheet, 默认位置为原最大下标加1.
	 * @author LFH
	 * @since 2017年5月19日 上午10:03:43
	 */
	public Sheet cloneSheet(int at) {
		return  this.cloneSheet(at,null);
	}

	/**
	 * <b>克隆一个sheet</b>
	 * @param at 原sheet位置
	 * @param sheetName 自定义名称(可选)
	 * @return 返回克隆成功的sheet, 默认位置为原最大下标加1.
	 * @author LFH
	 * @since 2017年5月19日 上午10:03:43
	 */
	public Sheet cloneSheet(int at, String sheetName) {
		Sheet sheet;
		if (this.theWorkbook == null || this.theWorkbook.getSheetAt(at) == null) {
			System.err.println("当前sheet尚为空,无法实现克隆操作!");
			return null;
		} else {
			sheet = this.theWorkbook.cloneSheet(at);
			if (sheetName != null && sheetName.length()>0&&sheetName.trim().length()>0) {
				this.theWorkbook.setSheetName(this.theWorkbook.getSheetIndex(sheet), sheetName.trim());
			}
			return sheet;
		}
	}

	/**
	 * <b>批量克隆指定sheet</b>
	 * @param at 原sheet位置
	 * @param times 复制次数
	 * @author LFH
	 * @since 2017年5月19日 上午11:59:39
	 */
	public void cloneSheet(int at, int times) {
		for (int i = 0; i < times; i++) {
			cloneSheet(at);
		}
	}

	/** <b>获取sheet下标</b>
	 * @param sheet 工作表对象
	 * @return 返回-1则说明此sheet无法获取.
	 * @author LFH
	 * @since 2017年5月19日 上午10:14:59
	 */
	public int getSheetIndex(Sheet sheet) {
		if (this.theWorkbook == null) {
			System.err.println("当前Workbook对象尚为空,无法实现操作!");
			return -1;
		}
		return this.theWorkbook.getSheetIndex(sheet);
	}

	/**
	 * <b>获取sheet下标</b>
	 * @param sheetName 工用表名称
	 * @return 返回-1则说明此sheet无法获取.
	 * @author LFH
	 * @since 2017年5月19日 上午10:14:57
	 */
	public int getSheetIndex(String sheetName) {
		if (this.theWorkbook == null) {
			System.err.println("当前Workbook对象尚为空,无法实现操作!");
			return -1;
		}
		return this.theWorkbook.getSheetIndex(sheetName);
	}

	/* ********************************************************************************************* */

	/**
	 * 批量填值
	 *
	 * @param list List[Map[String, String]]|| 数据集合
	 * @param rowk 行起始位置(0开始)
	 * @param cellk 列起始位置
	 * @param keys 键值数组[String]
	 * @author LFH
	 */
	public void entrySheet(List<Map<String, Object>> list, int rowk, int cellk, String[] keys) {
		this.entrySheet(list,rowk,cellk,keys,null);
	}

	/**
	 * 批量填值
	 *
	 * @param list List[Map[String, String]]|| 数据集合
	 * @param rowk 行起始位置(0开始)
	 * @param cellk 列起始位置
	 * @param keys 键值数组[String]
	 * @param type 类型集合 Map[Integer,Type(num,numz,date,formula,..)]
	 * @author LFH
	 */
	public void entrySheet(List<Map<String, Object>> list, int rowk, int cellk, String[] keys,
			Map<Integer, Type> type) {
		if (this.theSheet == null) {
			System.err.println("当前对象中sheet尚为空,无法实现填值操作!");
		} else {
			if (list != null && list.size() > 0) {
				if (rowk < 0) {
					System.err.println("请传入合法行标!(0-~)");
				} else if (cellk < 0) {
					System.err.println("请传入合法列标!(0-~)");
				} else if (keys == null || keys.length <= 0) {
					System.err.println("请传入合法键值数组!");
				} else {
					Row rowi ;
					Row rowf = this.entryRow(rowk);
					boolean hasTypesDefine = type!=null;
					for (int i = 0; i < list.size(); i++) {
						rowi = this.entryRow(i + rowk);
						if(hasTypesDefine){
							for (int j = 0; j < keys.length; j++) {
								this.entryCellInType(rowi, j + cellk, list.get(i).get(keys[j]), type.get(j + cellk));
							}
						}else{
							for (int j = 0; j < keys.length; j++) {
								this.entryCell(rowi, j + cellk, list.get(i).get(keys[j]));
							}
						}
						this.copyRowStyle(rowf, rowi, cellk, keys.length - 1 + cellk);
					}
				}
			} else {
				// System.err.println("数据为空,已终止填值!");
			}
		}
	}

	/**
	 * 批量填值(<b>对sheet 进行单行填值</b>)
	 *
	 * @param map Map[String, String]|| 单行数据Map集合
	 * @param rowk 操作行位置(0开始)
	 * @param cellk 列起始位置
	 * @param keys 键值数组[String]
	 * @author LFH
	 */
	public void entrySheetSingleRow(Map<String, Object> map, int rowk, int cellk, String[] keys) {
		this.entrySheetSingleRow(map,rowk,cellk,keys,null);
	}

	/**
	 * 批量填值(<b>对sheet 进行单行填值</b>)
	 *
	 * @param map Map[String, String]|| 单行数据Map集合
	 * @param rowk 操作行位置(0开始)
	 * @param cellk 列起始位置
	 * @param keys 键值数组[String]
	 * @param type 类型集合 Map[Integer,Type(num,numz,date,formula,..)](可选)
	 * @author LFH
	 */
	public void entrySheetSingleRow(Map<String, Object> map, int rowk, int cellk, String[] keys,
			Map<Integer, Type> type) {
		if (this.theSheet == null) {
			System.err.println("当前对象中sheet尚为空,无法实现填值操作!");
		} else {
			if (map != null && (!map.isEmpty())) {
				if (rowk < 0) {
					System.err.println("请传入合法行标!(0-~)");
				} else if (cellk < 0) {
					System.err.println("请传入合法列标!(0-~)");
				} else if (keys == null || keys.length <= 0) {
					System.err.println("请传入合法键值数组!");
				}  else {
					Row rowf = this.entryRow(rowk);
					boolean hasTypesDefine = type!=null;
					if (hasTypesDefine) {
						for (int j = 0; j < keys.length; j++) {
							this.entryCellInType(rowf, j + cellk, map.get(keys[j]), type.get(j + cellk));
						}
					}else{
						for (int j = 0; j < keys.length; j++) {
							this.entryCell(rowf, j + cellk, map.get(keys[j]));
						}
					}

				}
			} else {
				// System.err.println("数据为空,已终止填值!");
			}
		}
	}

	/**
	 * 不规则批量填充工作表
	 *
	 * @param positions 位置集合 <p>此参数由: {@link #cpsBatch(CellPosition...)} 方法构造<p/>
	 * @param map 数据集
	 * @author LFH
	 * @since 2018年3月7日 下午9:32:08
	 */
	public void entrySheet(List<CellPosition> positions, Map<String, Object> map) {
		if (positions != null && positions.size() > 0) {
			positions.forEach(i -> {
				entryCell(i.getRow(), i.getCell(), i.getKey() != null ? map.get(i.getKey()) : i.getValue(),
						i.getType());
			});
		}
	}

	/**
	 * 结合entrySheet导出Excel文档
	 *
	 * @param response 当前请求响应对象
	 * @author LFH
	 */
	public void outFile(HttpServletResponse response, String fname) {
		try {
			if (this.theWorkbook == null) {
				System.err.println("请先创建Sheet!");
			}
			this.outFile(response, fname, this.theWorkbook);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * 输出表格到硬盘指定位置
	 * @param outpath 目标路径
	 */
	public void outputWithExcel(String outpath) {
		try (FileOutputStream fos = new FileOutputStream(outpath)) {
			outputWithExcel(fos);
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
	/**
	 * 输出表格到输出流
	 * @param os 输出流
	 */
	public void outputWithExcel(OutputStream os){
		try(OutputStream theOs=os) {
			this.theWorkbook.write(theOs);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	/* ********************************************************************************************* */

	/**
	 * 获取有效单元格并设置值(可选择类型)
	 * !*!注意:<b>如果要设置日期类型数据,请传入Date类型数据,否则无法进行</b>
	 *
	 * @param row 行标
	 * @param index 列标
	 * @param value 值
	 * @param type 单元格类型
	 * @author LFH
	 */
	public void entryCell(int row, int index, Object value, Type type) {
		Row rowi = entryRow(row);
		switch (type) {
		case NORMAL:
			entryCell(rowi, index, value);
			break;
		case NUM:
			numCell(rowi, index, value);
			break;
		case NUMZ:
			numCellZ(rowi, index, value);
			break;
		case DATE:
			dateCell(rowi, index, (Date) value);
			break;
		case FORMULA:
			formulaCell(rowi, index, value);
			break;
		case RICH:
			/* 富文本 */
			entryCellRich(rowi, index, value);
			break;
		default:
			entryCell(rowi, index, value);
			break;
		}
	}

	/**
	 * 填充指定单元格
	 * @param row 行
	 * @param index 列
	 * @param value 值
	 */
	public void entryCell(int row, int index, Object value) {
		entryCell(this.entryRow(row), index, value);
	}

	/**
	 * 含样式的填值
	 *
	 * @param row 行
	 * @param index 列
	 * @param value 值
	 * @param cellStyle 样式
	 * @author LFH
	 * @since 2018年4月14日 下午10:34:40
	 */
	public void entryCell(int row, int index, Object value, CellStyle cellStyle) {
		entryCell(this.entryRow(row), index, value, cellStyle);
	}

	/**
	 * 获取有效合并单元格进行操作
	 *
	 * @param rs 开始行
	 * @param re 结束行
	 * @param s 开始列
	 * @param e 结束列
	 * @author LFH
	 */
	public void entryRegion(int rs, int re, int s, int e, Object v) {
		CellStyle style = this.getStyle(rs, s);
		CellRangeAddress region = new CellRangeAddress(rs, re, s, e);//
		this.theSheet.addMergedRegion(region);
		Row row = this.theSheet.getRow(rs);
		entryCell(row, s, v, style);
	}

	/**
	 * 获取有效合并单元格进行操作(数值类型)
	 *
	 * @param rs 开始行
	 * @param re 结束行
	 * @param s 开始列
	 * @param e 结束列
	 * @author LFH
	 */
	public void numRegion(int rs, int re, int s, int e, Object v) {
		CellRangeAddress region = new CellRangeAddress(rs, re, s, e);
		this.theSheet.addMergedRegion(region);
		Row row = this.theSheet.getRow(rs);
		numCell(row, s, v);
	}

	/**
	 * 获取有效单元格
	 *
	 * @param rowi 行
	 * @param celli 列
	 * @return Cell
	 * @author LFH
	 * @since 2017年11月3日 下午2:51:15
	 */
	public Cell entryCell(int rowi, int celli) {
		Row row = this.entryRow(rowi);
		Cell cell = row.getCell(celli);
		if (cell == null) {
			cell = row.createCell(celli);
		}
		return cell;
	}

	/**
	 * 获取有效行
	 *
	 * @param index 行
	 * @return Row
	 * @author LFH
	 * @since 2017年11月3日 下午2:15:34
	 */
	public Row entryRow(int index) {
		Row rowi = this.theSheet.getRow(index);
		if (rowi == null) {
			rowi = this.theSheet.createRow(index);
		}
		return rowi;
	}

	/**
	 * 对多个字符串值进行数值求和
	 * <b>不限整数或浮点数<b/>
	 *
	 * @param v1 必填参数1
	 * @param v2 可选参数列表2
	 * @return 和
	 * @author LFH
	 * @since 2017年5月12日 下午2:49:57
	 */
	public Object getSum(String v1, String... v2) {
		String a1 = v1;
		String[] a2 = v2;
		a1 = a1 == null || "".equals(a1) || "null".equalsIgnoreCase(a1) ? "0" : a1.trim();
		if (a2 != null && a2.length > 0) {
			for (int i = 0; i < a2.length; i++) {
				a2[i] = a2[i] == null || "".equals(a2[i]) || "null".equalsIgnoreCase(a2[i]) ? "0" : a2[i].trim();
			}
		}
		Object s = "";
		double x1 = 0.0, x2 = 0.0, xs = 0.0;
		int y1 = 0, y2 = 0, ys = 0;
		try {
			if (a1.indexOf(".") > 0) {
				x1 += Double.parseDouble(a1);
			} else {
				y1 += Integer.parseInt(a1);
			}
			for (String x : a2) {
				if (x.indexOf(".") > 0) {
					x2 += Double.parseDouble(x);
				} else {
					y2 += Integer.parseInt(x);
				}
			}
			xs = x1 + x2;
			ys = y1 + y2;
			if (xs > 0 && ys > 0) {
				s = (double) (xs + ys);
			} else if (xs > 0 && ys == 0) {
				s = (double) xs;
			} else {
				s = (int) ys;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return s;
	}

	/**
	 * 复制行样式(将行样式对应单元格进行复制)
	 *
	 * @param rowk 原行
	 * @param rowi 目标行
	 * @param s 开始列
	 * @param e 结束列
	 * @author LFH
	 */
	public void copyRowStyle(int rowk, int rowi, int s, int e) {
		Row _rowk = this.entryRow(rowk);
		Row _rowi = this.entryRow(rowi);
		this.copyRowStyle(_rowk, _rowi, s, e);
	}

	/**
	 * 设置指定行样式
	 * @param rowi 行标
	 * @param s 列开始
	 * @param e 列结束
	 * @param style 样式
	 * @author LFH
	 * @since 2017年5月17日 下午6:04:58
	 */
	public void setRowStyle(int rowi, int s, int e, CellStyle style) {
		if (this.theSheet == null) {
			System.err.println("请先创建Sheet!");
			return;
		}
		for (; s <= e; s++) {
			Cell cell = entryCell(entryRow(rowi), s);
			cell.setCellStyle(style);
		}
	}

	/**
	 * 设置指定行样式(批量设置)
	 * @param rowi 行标
	 * @param s 列开始
	 * @param e 列结束
	 * @param style 样式
	 * @param times 次数
	 * @author LFH
	 * @since 2017年5月17日 下午6:04:58
	 */
	public void setRowStyle(int rowi, int s, int e, CellStyle style, int times) {
		if (this.theSheet == null) {
			System.err.println("请先创建Sheet!");
			return;
		}
		for (int i = 0; i < times; i++) {
			setRowStyle(rowi + i, s, e, style);
		}
	}

	/**
	 * <b>创建CellStyle,以供额外的设置样式</b>
	 * @return CellStyle
	 * @author LFH
	 * @since 2017年5月17日 下午6:25:43
	 */
	public CellStyle createStyle() {
		if (this.theWorkbook == null) {
			System.err.println("请先创建Sheet!");
			return null;
		}
		return this.theCellStyle;
	}

	/**
	 * 创建数据格式
	 *
	 * @return DataFormat
	 * @author LFH
	 * @since 2018年3月2日 上午10:31:37
	 */
	public DataFormat createDataFormat() {
		if (this.theWorkbook == null) {
			System.err.println("请先创建Sheet!");
			return null;
		}
		return this.theWorkbook.createDataFormat();
	}

	/**
	 * 获取指定单元格CellStyle,以供额外的附加样式
	 * @param rowi 行标
	 * @param celli 列标
	 * @return CellStyle
	 * @author LFH
	 * @since 2017年5月17日 下午6:38:33
	 */
	public CellStyle getStyle(int rowi, int celli) {
		if (this.theSheet == null) {
			System.err.println("请先创建Sheet!");
			return null;
		}
		Cell cell = this.entryCell(this.entryRow(rowi), celli);
		return cell.getCellStyle();
	}

	/**
	 * 复制单元格
	 * @param fromRow 源行
	 * @param fromCell 源列
	 * @param toRow 目的行
	 * @param toCell 目的列
	 * @param copyValueFlag 是否包含内容
	 * @author LFH
	 * @since 2017年5月17日 下午1:52:49
	 */
	public void copyCell(int fromRow, int fromCell, int toRow, int toCell, boolean copyValueFlag) {
		if (this.theSheet == null) {
			System.err.println("请先创建Sheet!");
			return;
		}
		Cell s = entryCell(entryRow(fromRow), fromCell);
		Cell t = entryCell(entryRow(toRow), toCell);
		copyCell(s, t, copyValueFlag);
	}

	/**
	 * 复制行
	 * @param fromRow 源行
	 * @param toRow 目的行
	 * @param copyValueFlag 是否含内容复制
	 * @author LFH
	 */
	public void copyRow(int fromRow, int toRow, boolean copyValueFlag) {
		if (this.theSheet == null) {
			System.err.println("请先创建Sheet!");
			return;
		}
		Row f = entryRow(fromRow);
		Row t = entryRow(toRow);
		copyRow(f, t, copyValueFlag);
	}

	/**
	 * 复制行(指定范围列)
	 * @param fromRow 源行
	 * @param toRow 目的行
	 * @param fromCell 复制列的起始
	 * @param toCell 复制列的结束
	 * @param copyValueFlag 是否含内容复制
	 * @author LFH
	 * @since 2017年5月17日 下午2:01:07
	 */
	public void copyRow(int fromRow, int toRow, int fromCell, int toCell, boolean copyValueFlag) {
		for (int i = fromCell; i <= toCell; i++) {
			copyCell(fromRow, i, toRow, i, copyValueFlag);
		}
	}

	/**
	 * 清空行的值
	 * @param fromCell 起始列
	 * @param endCell 结束列
	 * @param rowk 待操作行
	 * @author LFH
	 * @since 2017年5月17日 下午4:58:56
	 */
	public void clearRow(int fromCell, int endCell, int... rowk) {
		if (this.theSheet == null) {
			System.err.println("请先创建Sheet!");
			return;
		}
		for (int i : rowk) {
			Row rowi = entryRow(i);
			for (; fromCell <= endCell; fromCell++) {
				Cell c = entryCell(rowi, fromCell);
				c.setCellType(CellType.BLANK);
			}
		}
	}

	/**
	 * 复制行(指定范围列)[可复制多次]
	 * @param fromRow 源行
	 * @param toRow 目的行
	 * @param times 复制次数(从目的行开始计算复制次数)
	 * @param fromCell 复制列的起始
	 * @param toCell 复制列的结束
	 * @param copyValueFlag 是否包含内容
	 * @author LFH
	 * @since 2017年5月17日 下午2:01:07
	 */
	public void copyRow(int fromRow, int toRow, int times, int fromCell, int toCell, boolean copyValueFlag) {
		this.copyRow(fromRow,toRow,times,fromCell,toCell,copyValueFlag,null);
	}

	/**
	 * 复制行(指定范围列)[可复制多次]
	 * @param fromRow 源行
	 * @param toRow 目的行
	 * @param times 复制次数(从目的行开始计算复制次数)
	 * @param fromCell 复制列的起始
	 * @param toCell 复制列的结束
	 * @param copyValueFlag 是否包含内容
	 * @param clear 是否清空内容(可选)
	 * @author LFH
	 * @since 2017年5月17日 下午2:01:07
	 */
	public void copyRow(int fromRow, int toRow, int times, int fromCell, int toCell, boolean copyValueFlag,
			Boolean clear) {
		for (int it = 0; it <= times; it++) {
			copyRow(fromRow, toRow++, fromCell, toCell, copyValueFlag);
			if (clear != null &&clear) {
				clearRow(fromCell, toCell, toRow - 1);
			}
		}
	}

	/**
	 * 移动行(携带样式和内容)
	 *
	 * @param startRow 移动区域起始行
	 * @param endRow 移动区域截止行
	 * @param n 移动几行
	 * @author LFH
	 * @since 2018年3月1日 上午10:07:45
	 */
	public void shiftRows(int startRow, int endRow, int n) {
		this.theSheet.shiftRows(startRow, endRow, n, true, true);
	}

	/* ------------------------------------------------------------------------ */
	/* 下方私有方法请勿设置公开! */

	/**
	 * 判断值是否为null
	 * @param value 值
	 * @return  boolean
	 */
	private static boolean isNull(Object value) {
		return value == null || "null".equals(value) || "".equals(value);
	}

	/**
	 * 填充时数据预处理
	 * @param value 值
	 * @param defaultValue 默认值("")
	 * @return 值
	 */
	private static String valuePrepare(Object value,String defaultValue){
		String v;
		if (isNull(value)) {
			v = Optional.ofNullable( defaultValue).orElse("");
		} else {
			v = value.toString();
		}
		return v;
	}

	/**
	 * 填充时数据预处理
	 * @param value 值
	 * @return 值
	 */
	private static String valuePrepare(Object value){
		return valuePrepare(value,null);
	}

	/**
	 * 复制行样式(将行样式对应单元格进行复制)
	 *
	 * @param rowk 原行
	 * @param rowi 目标行
	 * @param s 开始列
	 * @param e 结束列
	 * @author LFH
	 */
	private void copyRowStyle(Row rowk, Row rowi, int s, int e) {
		if (rowk == null) {
			return;
		}
		CellStyle style = this.createStyle();
		style.cloneStyleFrom(style);
		CellStyle rowStyle = rowk.getRowStyle();
		if (rowStyle != null) {
			rowi.setRowStyle(rowStyle);
		}
		Cell cell ;
		int i = s;
		for (; i <= e; i++) {
			style = entryCell(rowk, i).getCellStyle();
			cell = entryCell(rowi, i);
			cell.setCellStyle(style);
		}
	}

	/**
	 * 获得有效单元格
	 *
	 * @param rowi  行对象
	 * @param index 列
	 * @return cell
	 * @author LFH
	 */
	private Cell entryCell(Row rowi, int index) {
		Cell cell = rowi.getCell(index);
		if (cell == null) {
			cell = rowi.createCell(index);
		}
		return cell;
	}

	/**
	 * 获取有效单元格并设置值(含样式)
	 *
	 * @param rowi 行
	 * @param index 列标
	 * @param value 值
	 * @param style 样式
	 * @author LFH
	 * @since 2017年5月12日 下午2:47:04
	 */
	private void entryCell(Row rowi, int index, Object value, CellStyle style) {
		Cell cell = entryCell(rowi, index);
		cell.setCellStyle(style);
		String v =valuePrepare(value);
		cell.setCellValue(v);
	}

	/**
	 * 获取有效单元格并设置值(文本值)
	 *
	 * @param rowi 行
	 * @param index 列标
	 * @param value 值
	 * @author LFH
	 */
	private void entryCell(Row rowi, int index, Object value) {
		Cell cell = entryCell(rowi, index);
		String v = valuePrepare(value);
		cell.setCellValue(v);
	}

	/**
	 * 填充富文本值
	 * @param rowi 行
	 * @param index 列标
	 * @param value 值
	 * @author LFH
	 */
	private void entryCellRich(Row rowi, int index, Object value) {
		Cell cell = entryCell(rowi, index);
		String v = valuePrepare(value);
		try {
			cell.setCellValue(this.constructorDefine.richTextString.newInstance(v));
		}catch (IllegalAccessException|InstantiationException|InvocationTargetException e){
			e.printStackTrace();
		}

	}

	/**
	 * 获取有效单元格并设置值(数值类型);
	 * <b>非数值则为空</b>
	 * @param rowi 行
	 * @param index 列标
	 * @param value 值
	 * @author LFH
	 */
	private void numCell(Row rowi, int index, Object value) {
		Cell cell = entryCell(rowi, index);
		String v = valuePrepare(value,"0");
		if (v.contains(".")) {
			cell.setCellValue(Double.parseDouble(v));
		} else if (!"0".equals(v)) {
			cell.setCellValue(Integer.parseInt(v));
		} else {
			cell.setCellValue("");
		}
	}

	/**
	 * 获取有效单元格并设置值(数值类型);
	 * <b>非数值则为0</b>
	 * @param rowi 行
	 * @param index 列标
	 * @param value 值
	 * @author LFH
	 */
	private void numCellZ(Row rowi, int index, Object value) {
		Cell cell = entryCell(rowi, index);
		String v = valuePrepare(value,"0");
		if (v.contains(".")) {
			cell.setCellValue(Double.parseDouble(v));
		} else {
			cell.setCellValue(Integer.parseInt(v));
		}
	}

	/**
	 * 非数值或零值则为空.
	 *
	 * @param rowi 行
	 * @param index 列标
	 * @param value 值
	 * @author LFH
	 * @since 2017年12月8日 下午10:30:34
	 */
	private void numCellT(Row rowi, int index, Object value) {
		Cell cell = entryCell(rowi, index);
		String v = valuePrepare(value,"0");
		if (v.contains(".")) {
			cell.setCellValue(Double.parseDouble(v));
		} else if ("0".equals(v)) {
			cell.setCellValue("");
		} else {
			cell.setCellValue(Integer.parseInt(v));
		}
	}

	/**
	 * 日期单元格
	 *
	 * @param rowi 行
	 * @param index 列标
	 * @param value 值
	 * @author LFH
	 */
	private void dateCell(Row rowi, int index, Date value) {
		if (value==null) {
			System.err.println("请传入日期类型数据!");
			return;
		}
		Cell cell = entryCell(rowi, index);
		Date v = value;
		if (isNull(value)) {
			v = null;
		}
		cell.setCellValue(v);
	}

	/**
	 * 公式单元格
	 *
	 * @param rowi 行
	 * @param index 列标
	 * @param value 值
	 * @author LFH
	 */
	private void formulaCell(Row rowi, int index, Object value) {
		Cell cell = entryCell(rowi, index);
		String v = "";
		if (isNull(value)) {
			v = "";
			cell.setCellValue(v);
		} else {
			try {
				v = value.toString();
				cell.setCellFormula(v);
			} catch (Exception e) {
				cell.setCellValue(v);
				e.printStackTrace();
			}

		}
	}

	/**
	 * 获取web项目绝对路径.
	 *
	 * @param request 请求对象
	 * @return String
	 * @author LFH
	 * @since 2017年11月24日 上午12:56:39
	 */
	private static String getRootPath(HttpServletRequest request) {
		String rootPath = request.getSession().getServletContext().getRealPath("/");
		return rootPath.replaceAll("\\\\", "/");
	}

	/**
	 * 向浏览器发送文件流,以供下载
	 *
	 * @param response 响应对象
	 * @param fname    导出后名称
	 * @param wb 文档对象
	 * @author LFH
	 * @since 2017年5月12日 下午2:47:04
	 * @throws Exception 文件流异常
	 */
	private void outFile(HttpServletResponse response, String fname, Workbook wb) throws Exception {
		try (OutputStream out = response.getOutputStream()) {
			fname = fname == null || fname.trim().length() <= 0 ? "_" : fname;
			fname = new String(fname.getBytes("GBK"), StandardCharsets.ISO_8859_1);
			response.reset();
			response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8");
			// 靠这一行向外导出文件,("文件名以双引号包住,是为了避免在火狐等浏览器导出文件时文件名中有空格出现问题的情况 !")
			response.setHeader("Content-disposition", "attachment; filename=\"" + fname + ".xlsx" + "\"");
			wb.write(out);
		} catch (IOException e) {
			throw e;
		} finally {
			if (wb != null) {
				wb.close();
			}
		}
	}

	/**
	 * 获得文档流
	 * @param rootPath 文件根目录
	 * @param path 路径
	 * @return fis 文件流
	 * @author LFH
	 * @since 2017年5月12日 下午2:47:04
	 * @throws Exception 文件流获取异常
	 */
	private static FileInputStream fileInput(String rootPath, String path) throws Exception {
		String suffix=path.toUpperCase();
		if(!(suffix.endsWith("XLSX")||suffix.endsWith("XLS"))){
			throw new RuntimeException("请提供后缀为 .xls|.xlsx 格式的Excel 模板!");
		}
		String filePath  = rootPath + path;
		File file = new File(filePath);
		return new FileInputStream(file);
	}

	/**
	 * 私有方法用于批量填值用.
	 *
	 * @param rowi 行号
	 * @param c 列号
	 * @param v 值
	 * @param type 类型
	 * @author LFH
	 */
	private void entryCellInType(Row rowi, int c, Object v, Type type) {
		type = type == null ? Type.NORMAL : type;
		switch (type) {
		case NORMAL:
			entryCell(rowi, c, v);/* 常规 */
			break;
		case NUM:
			numCell(rowi, c, v);/* 非数值则空 */
			break;
		case NUMZ:
			numCellZ(rowi, c, v);/* 非数值则0 */
			break;
		case NUMT:
			numCellT(rowi, c, v);/* 非数值或0值则空 */
			break;
		case DATE:
			dateCell(rowi, c, (Date) v);/* 日期 */
			break;
		case FORMULA:
			formulaCell(rowi, c, v);/* 公式 */
			break;
		case RICH:
			entryCellRich(rowi, c, v);/* 富文本 */
			break;
		default:
			entryCell(rowi, c, v);
			break;
		}
	}

	/**
	 * 复制单元格
	 *
	 * @param srcCell 源单元格
	 * @param distCell 目标单元格
	 * @param copyValueFlag 是否含内容
	 * @author LFH
	 */
	private void copyCell(Cell srcCell, Cell distCell, boolean copyValueFlag) {
		// distCell.setEncoding(srcCell.getEncoding());
		// 目标单元格样式设置
		distCell.setCellStyle(srcCell.getCellStyle());
		//
		if (srcCell.getCellComment() != null) {
			distCell.setCellComment(srcCell.getCellComment());
		}
		// 单元格复制
		CellType srcCellType = srcCell.getCellTypeEnum();
		distCell.setCellType(srcCellType);
		if (copyValueFlag) {
			if (srcCellType == CellType.NUMERIC) {
				if (HSSFDateUtil.isCellDateFormatted(srcCell)) {
					distCell.setCellValue(srcCell.getDateCellValue());
				} else {
					distCell.setCellValue(srcCell.getNumericCellValue());
				}
			} else if (srcCellType == CellType.STRING) {
				distCell.setCellValue(srcCell.getRichStringCellValue());
			} else if (srcCellType == CellType.BLANK) {
				//
			} else if (srcCellType == CellType.BOOLEAN) {
				distCell.setCellValue(srcCell.getBooleanCellValue());
			} else if (srcCellType == CellType.ERROR) {
				distCell.setCellErrorValue(FormulaError.forInt(srcCell.getErrorCellValue()).getCode());
			} else if (srcCellType == CellType.FORMULA) {
				distCell.setCellFormula(srcCell.getCellFormula());
			} else { //
			}
		}
	}

	/**
	 * 复制行
	 * @param fromRow 源行
	 * @param toRow 目标行
	 * @param copyValueFlag 是否含内容复制
	 * @author LFH
	 */
	private void copyRow(Row fromRow, Row toRow, boolean copyValueFlag) {
		for (Iterator<Cell> cellIt = fromRow.cellIterator(); cellIt.hasNext();) {
			Cell tmpCell = cellIt.next();
			Cell newCell = toRow.createCell(tmpCell.getColumnIndex());
			copyCell(tmpCell, newCell, copyValueFlag);
		}
	}

	/**
	 * 主要为方便批量设置数据类型
	 * @param type 传入类型对照Map
	 * @param t 数据类型
	 * @param i 对应列号
	 * @return  Map<Integer, Type>
	 * @author LFH
	 * @since 2017年5月12日 下午5:52:38
	 */
	public Map<Integer, Type> mapType(Map<Integer, Type> type, Type t, Integer... i) {
		for (Integer index : i) {
			type.put(index, t);
		}
		return type;
	}

	/**
	 * 获取类型Map集合.
	 *
	 * @param t 数据类型
	 * @param i 对应列号
	 * @return Map<Integer, Type>
	 * @author LFH
	 * @since 2017年12月8日 下午10:19:40
	 */
	public Map<Integer, Type> mapType(Type t, Integer... i) {
		Map<Integer, Type> types = new HashMap<>();
		return i!=null&&i.length>0? mapType(types, t, i):types;
	}

	/**
	 * 获取类型Map集合.(静态方法)
	 *
	 * @param type Map<Integer, Type>
	 * @param t 数据类型
	 * @param i 对应列号
	 * @return  Map<Integer, Type>
	 * @author LFH
	 * @since 2018年3月20日 下午4:17:10
	 */
	public static Map<Integer, Type> mapTypes(Map<Integer, Type> type, Type t, Integer... i) {
		for (Integer index : i) {
			type.put(index, t);
		}
		return type;
	}

	/**
	 * 获取类型Map集合.(静态方法)
	 *
	 * @param t 数据类型
	 * @param i 对应列号
	 * @return  Map<Integer, Type>
	 * @author LFH
	 * @since 2018年3月20日 下午4:16:52
	 */
	public static Map<Integer, Type> mapTypes(Type t, Integer... i) {
		Map<Integer, Type> types = new HashMap<>();
		return i!=null&&i.length>0? mapTypes(types, t, i):types;
	}

	/**
	 * 快速创建位置对象
	 *
	 * @return  {@link CellPosition}
	 * @author LFH
	 * @since 2018年3月7日 下午9:24:16
	 */
	public CellPosition cps(String key, int row, int cell) {
		return new CellPosition(key, row, cell);
	}

	/**
	 * 快速创建位置对象
	 *
	 * @return {@link CellPosition}
	 * @author LFH
	 * @since 2018年3月11日 下午3:28:05
	 */
	public CellPosition cps(String key, int row, int cell, Type type) {
		return new CellPosition(key, row, cell, type);
	}

	/**
	 * 快速创建位置对象(附加值)
	 *
	 * @return {@link CellPosition}
	 * @author LFH
	 * @since 2018年3月11日 下午3:28:05
	 */
	public CellPosition cps(int row, int cell, Object value) {
		return new CellPosition(row, cell, value);
	}

	/**
	 * 快速创建位置对象(附加值)
	 *
	 * @return {@link CellPosition}
	 * @author LFH
	 * @since 2018年3月11日 下午3:28:05
	 */
	public CellPosition cps(int row, int cell, Object value, Type type) {
		return new CellPosition(row, cell, value, type);
	}

	/**
	 * 快速创建批量位置集合
	 *
	 * @param p 位置对象 <div>
	 *       <h3>对象参数p 由 {@link #cps(String, int, int)},{@link #cps(String, int, int, Type)} 方法创建</h3>
	 *          <div/>
	 * @return {@link CellPosition}
	 * @author LFH
	 * @since 2018年3月7日 下午9:28:11
	 */
	public List<CellPosition> cpsBatch(CellPosition... p) {
		return new ArrayList<>(Arrays.asList(p));
	}

	/* 单元格简单设置操作工具类 */

	/**
	 * 创建一个配置样式的对象.
	 * 注意:<b>此设置应尽量在批量填值完成后操作</b>.
	 * <div>
	 *     <p><b>example:</b></p>
	 *     <div>
	 *       CellStyle style = exp.getStyle(6, 1);<br/>
	 * 	          exp.entryCell(6, 0).setCellStyle( exp.createExpStyle(style)<br/>
	 * 	        .setFontColor(Color.RED.index)<br/>
	 * 	        .setFgColor(Color.YELLOW.index)<br/>
	 * 	        .setFontSize(20).setAlign("right", "bottom").finish());
	 *     </div>
	 * </div>
	 *
	 * @param expStyle 样式
	 * @return {@link ExpStyle}
	 * @author LFH
	 * @since 2018年4月14日 下午9:09:11
	 */
	public ExpStyle createExpStyle(CellStyle expStyle) {
		return new ExpStyle(this.theWorkbook, expStyle);
	}

	/**
	 * @return {@link ExpStyle}
	 * @author LFH
	 * @since 2018年4月14日 下午10:21:58
	 * @see #createExpStyle(CellStyle)
	 */
	public ExpStyle createExpStyle() {
		return new ExpStyle(this.theWorkbook);
	}

	/**
	 *  在表格中插入图片
	 * 目前支持<b>jpg/jpeg/gif/png/bmp</b>格式图片
	 * @param option 图片参数对象
	 * @author LFH
	 * @throws Exception
	 */
	public void insertImg(ImgOption option) throws Exception {
		String type = "";
		String[] types = { "jpg", "png", "gif", "bmp", "jpeg" };
		boolean check = false;
		String fPath = option.getfPath();
		if (fPath != null && fPath.contains(".")) {
			type = fPath.substring(fPath.lastIndexOf(".") + 1, fPath.length());
			type = type.trim();
			for (String t : types) {
				if (t.equalsIgnoreCase(type)) {
					check = true;
				}
			}
			if (!check) {
				throw new RuntimeException("Image's type is not defined or the type can't be read!\t" + type);
			}
		} else {
			throw new RuntimeException("Image File Not Found Exception!\t" + type);
		}
		// 新建图片缓存区
		BufferedImage bufferImg = null;
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		bufferImg = ImageIO.read(new File(this.rootPath + fPath));
		ImageIO.write(bufferImg, type, byteArrayOut);
		Drawing patriarch = this.theSheet.createDrawingPatriarch();
		// anchor对象创建
		ClientAnchor anchor = this.constructorDefine.clientAnchor.newInstance (option.getDx1(), option.getDy1(), option.getDx2(),
				option.getDy2(), option.getCol1(), option.getRow1(), option.getCol2(), option.getRow2());
		//		anchor.setAnchorType(3); // 3.15用法
		anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);//3.16用法
		// 插入图片
		patriarch.createPicture(anchor,
				this.theWorkbook.addPicture(byteArrayOut.toByteArray(), Workbook.PICTURE_TYPE_PNG));
	}

	/**
	 * 图片属性
	 */
	public ImgOption initImgOption(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2,
			String path) {
		return new ImgOption(dx1, dy1, dx2, dy2, (short) col1, row1, (short) col2, row2, path);
	}
	/**
	 * 图片属性
	 */
	public ImgOption initImgOption(int dx2, int dy2, int col1, int row1, int col2, int row2, String path) {
		return new ImgOption(dx2, dy2, (short) col1, row1, (short) col2, row2, path);
	}
	/**
	 * 图片属性
	 */
	public ImgOption initImgOption(int col1, int row1, int col2, int row2, String path) {
		return new ImgOption((short) col1, row1, (short) col2, row2, path);
	}

	/* 类与枚举定义 */

	/**
	 * 支持类型
	 */
	public enum WorkType{
		XLS(HSSFWorkbook.class),XLSX(XSSFWorkbook.class);
		private Class<? extends Workbook> workClass;

		public Class<? extends Workbook> getWorkClass() {
			return workClass;
		}

		/*通过此枚举类型去创建ExportExcelPlus工具实例*/

		/**
		 * @see ExportExcelPlus#createExportWork(WorkType, HttpServletRequest, String)
		 * */
		public ExportExcelPlus build(HttpServletRequest request,String path){
			return ExportExcelPlus.createExportWork(this,request,path);
		}

		/**
		 * @see ExportExcelPlus#createExportWork(WorkType, HttpServletRequest, String,int)
		 * */
		public ExportExcelPlus build(HttpServletRequest request,String path,int at){
			return ExportExcelPlus.createExportWork(this,request,path,at);
		}

		/**
		 * @see ExportExcelPlus#createExportWork(WorkType, String, String,int)
		 * */
		public ExportExcelPlus build(String rootPath,String path,int at){
			return ExportExcelPlus.createExportWork(this,rootPath,path,at);
		}
		/**
		 * @see ExportExcelPlus#createExportWork(WorkType, String, String)
		 * */
		public ExportExcelPlus build(String rootPath,String path){
			return ExportExcelPlus.createExportWork(this,rootPath,path,0);
		}

		WorkType(Class<? extends Workbook> workClass) {
			this.workClass = workClass;
		}
	}

	/**
	 * 一些用到的构造器类
	 */
	private class ConstructorDefine{
		private Constructor<? extends RichTextString> richTextString;
		private Constructor<? extends ClientAnchor> clientAnchor;

		ConstructorDefine(Constructor<? extends RichTextString> richTextStringConstructor,
				Constructor<? extends ClientAnchor> clientAnchorConstructor) {
			this.richTextString = richTextStringConstructor;
			this.clientAnchor = clientAnchorConstructor;
		}
	}


	/**
	 * 自定义的填充数据类型
	 * @author LFH
	 * @since 2017年5月12日
	 */
	public enum Type {
		NUM, NUMZ, NUMT, DATE, FORMULA, NORMAL, RICH

	}


	/**
	 * 填充sheet时用到的位置对象.
	 *
	 * @author LFH
	 * @since 2018年3月7日
	 */
	public class CellPosition {
		private String key;// 键
		private int row;// 行
		private int cell;// 列
		private Object value;
		private Type type = Type.NORMAL;// 类型

		private String getKey() {
			return key;
		}

		private int getRow() {
			return row;
		}

		private int getCell() {
			return cell;
		}

		private Type getType() {
			return type;
		}

		public Object getValue() {
			return value;
		}

		private CellPosition(String key, int row, int cell, Type type) {
			this.key = key;
			this.row = row;
			this.cell = cell;
			this.type = type;
		}

		private CellPosition(String key, int row, int cell) {
			this.key = key;
			this.row = row;
			this.cell = cell;
		}

		private CellPosition(int row, int cell, Object value) {
			this.row = row;
			this.cell = cell;
			this.value = value;
		}

		private CellPosition(int row, int cell, Object value, Type type) {
			this.row = row;
			this.cell = cell;
			this.value = value;
			this.type = type;
		}

	}

	/**
	 * 图片属性对象
	 */
	private class ImgOption {
		private int dx1 = 0;
		private int dy1 = 0;
		private int dx2 = 255;
		private int dy2 = 255;
		private short col1;
		private int row1;
		private short col2;
		private int row2;
		private String fPath;

		 String getfPath() {
			return fPath;
		}

		 int getDx1() {
			return dx1;
		}

		 int getDy1() {
			return dy1;
		}

		 int getDx2() {
			return dx2;
		}

		 int getDy2() {
			return dy2;
		}

		 short getCol1() {
			return col1;
		}

		 int getRow1() {
			return row1;
		}

		 short getCol2() {
			return col2;
		}

		 int getRow2() {
			return row2;
		}

		private ImgOption(int dx1, int dy1, int dx2, int dy2, short col1, int row1, short col2, int row2, String path) {
			this(col1,row1,col2,row2,path);
			this.dx1 = dx1;
			this.dy1 = dy1 > 255 ? 255 : dy1;
			this.dx2 = dx2;
			this.dy2 = dy2 > 255 ? 255 : dy2;
		}

		private ImgOption(int dx2, int dy2, short col1, int row1, short col2, int row2, String path) {
			this(col1,row1,col2,row2,path);
			this.dx2 = dx2;
			this.dy2 = dy2;
		}

		private ImgOption(short col1, int row1, short col2, int row2, String path) {
			this.col1 = col1;
			this.row1 = row1;
			this.col2 = col2;
			this.row2 = row2;
			this.fPath = path;
		}

	}

	/**
	 * 对齐方式变量-水平
	 *
	 * @author LFH
	 * @date 2018年4月14日
	 */
	private enum TAlign {
		center(HorizontalAlignment.CENTER), left(HorizontalAlignment.LEFT), right(HorizontalAlignment.RIGHT);
		private HorizontalAlignment ment;

		 TAlign(HorizontalAlignment ment) {
			this.ment = ment;
		}
	}

	/**
	 * 对齐方式变量-垂直
	 *
	 * @author LFH
	 * @date 2018年4月14日
	 */
	private enum VAlign {
		middle(VerticalAlignment.CENTER), top(VerticalAlignment.TOP), bottom(VerticalAlignment.BOTTOM);
		private VerticalAlignment ment;

		 VAlign(VerticalAlignment ment) {
			this.ment = ment;
		}
	}

	/**
	 * 样式配置类
	 *
	 * @author LFH
	 * @since 2018年4月14日
	 */
	public class ExpStyle {
		private CellStyle expStyle;
		private Font expFont;
		private boolean hasFont = false;

		private ExpStyle(Workbook work) {
			this.expStyle = work.createCellStyle();
			this.expFont = work.createFont();
		}

		/**
		 * 复制对象可读写属性.
		 *
		 * @author LFH
		 * @since 2018年4月14日
		 */
		private BinaryOperator<Object> copy = (x, y) -> {
			Class<?> type = x.getClass();
			try {
				BeanInfo beanInfo = Introspector.getBeanInfo(type);
				PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();// 获取属性数组
				for (PropertyDescriptor pd : propertyDescriptors) {
					String propertyName = pd.getName();// 获取属性名
					if (!"class".equalsIgnoreCase(propertyName)) {
						Method get = pd.getReadMethod();// 得到读属性方法(get...())
						Method set = pd.getWriteMethod();
						if (get != null) {
							Object value = get.invoke(x);// 获取属性值
							if (value != null && set != null) {
								set.invoke(y, value);
							}
						}
					}
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
			return y;
		};

		private ExpStyle(Workbook work, CellStyle expStyle) {
			this.expStyle = work.getNumCellStyles() < 3500 ? work.createCellStyle()
					: work.getCellStyleAt((short) (work.getNumCellStyles() - 1));
			copy.apply(expStyle, this.expStyle);
			this.expFont = work.createFont();
		}

		private void addFont() {
			if (!this.hasFont) {
				this.expStyle.setFont(this.expFont);
				this.hasFont = true;
			}
		}

		/**
		 * 添加字体颜色
		 *
		 * @param color {@link Color} 的<b>颜色属性</b> 的 <b>index</b> 属性
		 * @return {@link ExpStyle }
		 * @author LFH
		 * @since 2018年4月14日 下午8:48:22
		 */
		public ExpStyle setFontColor(short color) {
			this.expFont.setColor(color);
			addFont();
			return this;
		}

		/**
		 * 添加字体大小
		 *
		 * @param size FontSize
		 * @return {@link ExpStyle }
		 * @author LFH
		 * @since 2018年4月14日 下午8:56:36
		 */
		public ExpStyle setFontSize(int size) {
			this.expFont.setFontHeightInPoints((short) size);
			addFont();
			return this;
		}

		/**
		 * 设置单元格对齐
		 *
		 * @param tAlign 水平
		 * @param vAlign 垂直
		 * @return {@link ExpStyle }
		 * @author LFH
		 * @since 2018年4月14日 下午8:59:36
		 */
		public ExpStyle setAlign(String tAlign, String vAlign) {
			try {
				TAlign talign = TAlign.valueOf(tAlign);
				this.expStyle.setAlignment(talign.ment);
			} catch (Exception e) {
				//TODO
			}
			try {
				VAlign valign = VAlign.valueOf(vAlign);
				this.expStyle.setVerticalAlignment(valign.ment);
			} catch (Exception e) {
				//TODO
			}
			return this;
		}

		/**
		 * 设置背景颜色
		 *
		 * @param color {@link Color} 的<b>颜色属性</b> 的 <b>index</b> 属性
		 * @return {@link ExpStyle }
		 * @author LFH
		 * @since 2018年4月14日 下午9:11:59
		 */
		public ExpStyle setBgColor(short color) {
			this.expStyle.setFillBackgroundColor(color);
			this.expStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			return this;
		}

		/**
		 * 设置前景颜色
		 *
		 * @param color {@link Color} 的<b>颜色属性</b> 的 <b>index</b> 属性
		 * @return {@link ExpStyle }
		 * @author LFH
		 * @since 2018年4月14日 下午9:11:59
		 */
		public ExpStyle setFgColor(short color) {
			this.expStyle.setFillForegroundColor(color);
			this.expStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			return this;
		}

		/**
		 * 完成设置,返回样式.
		 *
		 * @return CellStyle
		 * @author LFH
		 * @since 2018年4月14日 下午9:59:53
		 */
		public CellStyle finish() {
			return this.expStyle;
		}
	}

}
