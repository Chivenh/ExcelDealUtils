package com.flytiger.Excel.v315;

import java.awt.image.BufferedImage;
import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.util.*;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

/**
 *  @since 2019/3/3 16:21
 *  废弃;请转向:{@link ExportExcelPlus}
 * @author LFH
 * @since  2017年5月12日
 * @see POI导表工具类(使用POI[3.15])
 * @see <b>.xlsx文件格式.</b>
 * @see 此类中方法基本已全部修改为对象级方法,即先有初始化,相关方法才能起效
 * @see --添加内容时请同时添加详细注释
 * @version 1.0.0
 * @see ***** 一般使用顺序
 * @see #createExportWork(HttpServletRequest, String, int)
 * @see --初始化
 * @see #entrySheet(List, int, int, String[])
 * @see --填充值
 * @see #outFile(HttpServletResponse, String)
 * @see --最后导出文件
 * @see ******详细参数请看对应方法注释
			 
																																	 
			
 * @see ******其它对象方法请参看公共方法列表
 *************/
@Deprecated
 public class ExportExcelXUtil {
	/************************************************************************************************/
	private XSSFSheet tSheet;// 当前操作sheet
	private XSSFWorkbook tWorkbook;// 当前操作workbook(不能get获取,保证对象结构完善)
	private String rootPath;// web项目根路径.
	private XSSFCellStyle xssfCellStyle;// 只创建一次样式进行复用,以免超过循环创建4000报错,

	private ExportExcelXUtil() {

	}

	private ExportExcelXUtil(XSSFSheet tSheet, XSSFWorkbook tWorkbook, String rootPath) {
		this.tSheet = tSheet;
		this.tWorkbook = tWorkbook;
		this.rootPath = rootPath;
		this.xssfCellStyle = this.tWorkbook.createCellStyle();
	}
	
	/**
	 * 初始化sheet
	 * 
	 * @author LFH
	 * @param request
	 * @param path Excel文件路径(注意:此path为相对于项目根目录的路径)
	 * @param at 工作表序号(0开始)
	 * @see <b> path 中请勿拼接项目根目录,此工具中会自动去获取项目根目录</b>
	 * @see 若模板保存在非项目所在目录中,请参看--
	 * @see #createExportWork(String, String, int)
	 */
	public static ExportExcelXUtil createExportWork(HttpServletRequest request, String path, int at) {
		return createExportWork(getRootPath(request), path, at);
	}

	/**
	 * 此初始化方法针对于模板并非在项目目录下的情况
	 * 
	 * @author LFH
	 * @date 2018年3月31日 下午3:51:45
	 * @param rootPath 模板在电脑硬盘中的目录
	 * @param path 模板路径,此路径是模板在rootPath目录之后的路径
	 * @param at 工作表序号(0开始)
	 * @see --[要获取多个sheet时,只要第一次初始化了工作对象,后面可以切换和调用]
	 * @see #settSheet(int)
	 * @see --来切换当前sheet
	 * @return
	 */
	public static ExportExcelXUtil createExportWork(String rootPath, String path, int at) {

		FileInputStream fis = null;
		XSSFSheet tSheet = null;
		XSSFWorkbook wb = null;
		ExportExcelXUtil exp = null;
		try {
			fis = fileInput(rootPath, path);
			if (fis != null) {
				wb = new XSSFWorkbook(fis);
				tSheet = wb.getSheetAt(at);
				exp = new ExportExcelXUtil(tSheet, wb, rootPath);
			} else {
				throw new RuntimeException("文件流未获取,请检查路径!");
			}

		} catch (Exception e) {
			// TODO Auto-generated catch block
			System.err.println(e);
		}
		return exp;
	}
	
	/**
	 * 设置文档unwriteProtectWorkbook,让文档取消只读.
	 * 
	 * @author LFH
	 * @date 2018年5月8日 下午4:48:20
	 * @return
	 */
	public ExportExcelXUtil setSheetWriteable() {
		if (this.tWorkbook != null) {
			this.tWorkbook.unLock();
		}
		return this;
	}

	/**
	 * @author LFH
	 * @date 2017年5月19日 上午10:57:52
	 * @see 切换当前sheet,前提是已经创建了模板对象,即已经调用过
	 * @see #settSheet(HttpServletRequest, String, int)
	 * @see --或者已克隆生成新sheet
	 * @see #cloneSheet(int, String)
	 * @param at 下标
	 */
	public void settSheet(int at) {
		if (this.tWorkbook == null || this.tWorkbook.getSheetAt(at) == null) {
			System.err.println("当前sheet尚为空,无法实现切换操作!\n请先初始化sheet,或克隆已存在的sheet");
			return;
		}
		XSSFSheet sheet = this.tWorkbook.getSheetAt(at);
		if (sheet == null) {
			throw new RuntimeException("工作表获取异常,请检查下标!");
		} else {
			this.tSheet = sheet;
		}
	}

	/**
	 * @author LFH
	 * @date 2017年5月19日 上午11:18:27
	 * @see 设置sheet名称
	 * @param at 下标
	 * @param name 名称
	 */
	public void setSheetName(int at, String name) {
		if (this.tWorkbook == null || this.tWorkbook.getSheetAt(at) == null) {
			System.err.println("当前sheet尚为空,无法实现操作");
			return;
		} else if (name == null || name.trim().length() < 1) {
			return;
		}
		this.tWorkbook.setSheetName(at, name);
	}

	/**
	 * @author LFH
	 * @date 2017年5月12日 下午5:34:47
	 * @see 主要用于当表格复杂时,使用此方法获取sheet进行填充
	 * @return
	 */
	public XSSFSheet gettSheet() {
		return tSheet;
	}

	/**
	 * 保护指定工作表,如果不指定,默认index为当前sheet.
	 * 
	 * @author LFH
	 * @date 2017年11月28日 下午9:54:47
	 * @param password
	 * @param index
	 * @return
	 */
	public boolean protectSheet(String password, int... index) {
		boolean b = false;
		try {
			if (index != null && index.length > 0) {
				for (int i : index) {
					XSSFSheet xSheet = this.tWorkbook.getSheetAt(i);
					if (xSheet != null) {
						xSheet.protectSheet(password);
					}
				}
			} else {
				this.tSheet.protectSheet(password);
			}
			b = true;
		} catch (Exception e) {
			b = false;
			throw (e);
		}
		return b;
	}

	/**
	 * @author LFH
	 * @date 2017年5月19日 上午10:03:43
	 * @see 克隆一个sheet
	 * @param at 原sheet位置
	 * @param sheetName 自定义名称(可选)
	 * @return 返回克隆成功的sheet,默认位置为原最大下标加1.
	 */
	public XSSFSheet cloneSheet(int at, String... sheetName) {
		XSSFSheet sheet = null;
		if (this.tWorkbook == null || this.tWorkbook.getSheetAt(at) == null) {
			System.err.println("当前sheet尚为空,无法实现克隆操作!");
			return null;
		} else {
			sheet = this.tWorkbook.cloneSheet(at);
			if (sheetName != null && sheetName.length > 0 && sheetName[0].trim().length() > 0) {
				this.tWorkbook.setSheetName(this.tWorkbook.getSheetIndex(sheet), sheetName[0].trim());
			}
			return sheet;
		}
	}

	/**
	 * @author LFH
	 * @date 2017年5月19日 上午11:59:39
	 * @see 批量克隆指定sheet
	 * @param at 原sheet位置
	 * @param times 复制次数
	 */
	public void cloneSheet(int at, int times) {
		for (int i = 0; i < times; i++) {
			cloneSheet(at);
		}
	}

	/**
	 * @author LFH
	 * @date 2017年5月19日 上午10:14:59
	 * @see 获取sheet下标
	 * @param sheet
	 * @return 返回-1则说明此sheet无法获取.
	 */
	public int getSheetIndex(XSSFSheet sheet) {
		if (this.tWorkbook == null) {
			System.err.println("当前Workbook对象尚为空,无法实现操作!");
			return -1;
		}
		return this.tWorkbook.getSheetIndex(sheet);
	}

	/**
	 * @author LFH
	 * @date 2017年5月19日 上午10:14:57
	 * @see 获取sheet下标
	 * @param sheetName
	 * @return 返回-1则说明此sheet无法获取.
	 */
	public int getSheetIndex(String sheetName) {
		if (this.tWorkbook == null) {
			System.err.println("当前Workbook对象尚为空,无法实现操作!");
			return -1;
		}
		return this.tWorkbook.getSheetIndex(sheetName);
	}

	/************************************************************************************************/

	/**
	 * 批量填值
	 * 
	 * @author LFH
	 * @param list List[Map[String, String]]|| 数据集合
	 * @param rowk 行起始位置(0开始)
	 * @param cellk 列起始位置
	 * @param keys 键值数组[String]
	 */
	public void entrySheet(List<Map<String, Object>> list, int rowk, int cellk, String[] keys) {
		if (this.tSheet == null) {
			System.err.println("当前对象中sheet尚为空,无法实现填值操作!");
			return;
		} else {
			if (list != null && list.size() > 0) {
				if (rowk < 0) {
					System.err.println("请传入合法行标!(0-~)");
					return;
				} else if (cellk < 0) {
					System.err.println("请传入合法列标!(0-~)");
					return;
				} else if (keys == null || keys.length <= 0) {
					System.err.println("请传入合法键值数组!");
					return;
				} else {
					XSSFRow rowi = null;
					XSSFRow rowf = this.entryRow(rowk);
					for (int i = 0; i < list.size(); i++) {
						rowi = this.entryRow(i + rowk);
						for (int j = 0; j < keys.length; j++) {
							this.entryCell(rowi, j + cellk, list.get(i).get(keys[j]));
						}
						this.copyRowStyle(rowf, rowi, cellk, keys.length - 1 + cellk);
					}
				}
			} else {
				// System.err.println("数据为空,已终止填值!");
				return;
			}
		}
	}

	/**
	 * 批量填值
	 * 
	 * @author LFH
	 * @param list List[Map[String, String]]|| 数据集合
	 * @param rowk 行起始位置(0开始)
	 * @param cellk 列起始位置
	 * @param keys 键值数组[String]
	 * @param type 类型集合 Map[Integer,Type(num,numz,date,formula,..)]
	 */
	public void entrySheet(List<Map<String, Object>> list, int rowk, int cellk, String[] keys,
			Map<Integer, Type> type) {
		if (this.tSheet == null) {
			System.err.println("当前对象中sheet尚为空,无法实现填值操作!");
			return;
		} else {
			if (list != null && list.size() > 0) {
				if (rowk < 0) {
					System.err.println("请传入合法行标!(0-~)");
					return;
				} else if (cellk < 0) {
					System.err.println("请传入合法列标!(0-~)");
					return;
				} else if (keys == null || keys.length <= 0) {
					System.err.println("请传入合法键值数组!");
					return;
				} else if (type == null) {
					System.err.println("请传入合法类型Map!");
					return;
				} else {
					XSSFRow rowi = null;
					XSSFRow rowf = this.entryRow(rowk);
					for (int i = 0; i < list.size(); i++) {
						rowi = this.entryRow(i + rowk);
						for (int j = 0; j < keys.length; j++) {
							if (type != null) {
								this.entryCellInType(rowi, j + cellk, list.get(i).get(keys[j]), type.get(j + cellk));
							} else {
								this.entryCell(rowi, j + cellk, list.get(i).get(keys[j]));
							}
						}
						this.copyRowStyle(rowf, rowi, cellk, keys.length - 1 + cellk);
					}
				}
			} else {
				// System.err.println("数据为空,已终止填值!");
				return;
			}
		}
	}

	/**
	 * 批量填值
	 * 
	 * @author LFH
	 * @see 对sheet 进行单行填值
	 * @param map Map[String, String]|| 单行数据Map集合
	 * @param rowk 操作行位置(0开始)
	 * @param cellk 列起始位置
	 * @param keys 键值数组[String]
	 * @param type 类型集合 Map[Integer,Type(num,numz,date,formula,..)](可选)
	 */
	@SuppressWarnings("unchecked")
	public void entrySheetSingleRow(Map<String, Object> map, int rowk, int cellk, String[] keys,
			Map<Integer, Type>... type) {
		if (this.tSheet == null) {
			System.err.println("当前对象中sheet尚为空,无法实现填值操作!");
			return;
		} else {
			if (map != null && (!map.isEmpty())) {
				if (rowk < 0) {
					System.err.println("请传入合法行标!(0-~)");
					return;
				} else if (cellk < 0) {
					System.err.println("请传入合法列标!(0-~)");
					return;
				} else if (keys == null || keys.length <= 0) {
					System.err.println("请传入合法键值数组!");
					return;
				} else if (type.length > 0 && type[0] == null) {
					System.err.println("请传入合法类型Map!");
					return;
				} else {
					XSSFRow rowf = this.entryRow(rowk);
					for (int j = 0; j < keys.length; j++) {
						if (type != null && type.length > 0 && type[0] != null) {
							this.entryCellInType(rowf, j + cellk, map.get(keys[j]), type[0].get(j + cellk));
						} else {
							this.entryCell(rowf, j + cellk, map.get(keys[j]));
						}
					}
				}
			} else {
				// System.err.println("数据为空,已终止填值!");
				return;
			}
		}
	}

	/**
	 * 不规则批量填充工作表
	 * 
	 * @see -- 第一个参数由
	 * @see #cpsBatch(CellPosition...)
	 * @see -- 方法构造
	 * @author LFH
	 * @date 2018年3月7日 下午9:32:08
	 * @param positions
	 * @param map
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
	 * @author LFH
	 * @param response
	 */
	public void outFile(HttpServletResponse response, String fname) {
		try {
			if (this.tWorkbook == null) {
				System.err.println("请先创建Sheet!");
			}
			this.outFile(response, fname, this.tWorkbook);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/************************************************************************************************/

	/**
	 * 获取有效单元格并设置值(可选择类型)
	 * !*!注意:如果要设置日期类型数据,请传入Date类型数据,否则无法进行
	 * 
	 * @author LFH
	 * @param row 行标
	 * @param index 列标
	 * @param value 值
	 * @param type 单元格类型
	 * @return
	 */
	public void entryCell(int row, int index, Object value, Type type) {
		XSSFRow rowi = entryRow(row);
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
			entryCellRich(rowi, index, value);/** 富文本 */
			break;
		default:
			entryCell(rowi, index, value);
			break;
		}
	}


	/**
	 * 无样式填值
	 * 
	 * @author LFH
	 * @date 2018年4月14日 下午10:35:14
	 * @param row
	 * @param index
	 * @param value
	 */
	public void entryCell(int row, int index, Object value) {
		entryCell(this.entryRow(row), index, value);
	}

	/**
	 * 含样式的填值
	 * 
	 * @author LFH
	 * @date 2018年4月14日 下午10:34:40
	 * @param row
	 * @param index
	 * @param value
	 * @param cellStyle
	 */
	public void entryCell(int row, int index, Object value, XSSFCellStyle cellStyle) {
		entryCell(this.entryRow(row), index, value, cellStyle);
	}

	/**
	 * 获取有效合并单元格进行操作
	 * 
	 * @author LFH
	 * @param sheet
	 * @param rs 开始行
	 * @param re 结束行
	 * @param s 开始列
	 * @param e 结束列
	 * @return
	 */
	public void entryRegion(int rs, int re, int s, int e, Object v) {
		XSSFCellStyle style = this.getStyle(rs, s);
		CellRangeAddress region = new CellRangeAddress(rs, re, s, e);//
		this.tSheet.addMergedRegion(region);
		XSSFRow row = this.tSheet.getRow(rs);
		entryCell(row, s, v, style);
	}

	/**
	 * 获取有效合并单元格进行操作(数值类型)
	 * 
	 * @author LFH
	 * @param sheet
	 * @param rs 开始行
	 * @param re 结束行
	 * @param s 开始列
	 * @param e 结束列
	 * @return
	 */
	public void numRegion(int rs, int re, int s, int e, Object v) {
		CellRangeAddress region = new CellRangeAddress(rs, re, s, e);
		this.tSheet.addMergedRegion(region);
		XSSFRow row = this.tSheet.getRow(rs);
		numCell(row, s, v);
	}
	
	/**
	 * 获取有效行
	 * 
	 * @author LFH
	 * @date 2017年11月3日 下午2:15:34
	 * @param index
	 * @return
	 */
	public XSSFRow entryRow(int index) {
		XSSFRow rowi = this.tSheet.getRow(index);
		if (rowi == null) {
			rowi = this.tSheet.createRow(index);
		}
		return rowi;
	}

	/**
	 * 获取有效单元格
	 * 
	 * @author LFH
	 * @date 2017年11月3日 下午2:51:15
	 * @param rowi
	 * @param celli
	 * @return
	 */
	public XSSFCell entryCell(int rowi, int celli) {
		XSSFRow row = this.entryRow(rowi);
		XSSFCell cell = row.getCell(celli);
		if (cell == null) {
			cell = row.createCell(celli);
		}
		return cell;
	}



	/**
	 * 对多个字符串值进行数值求和
	 * 
	 * @author LFH
	 * @date 2017年5月12日 下午2:49:57
	 * @see 不限整数或浮点数
	 * @param v1 必填参数1
	 * @param v2 可选参数列表2
	 * @return 和
	 */
	public Object getSum(String v1, String... v2) {
		String a1 = v1;
		String a2[] = v2;
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
			System.err.println(a1);
			System.err.println(a2.toString());
			System.err.println(e);
		}
		return s;
	}


	/**
	 * 复制行样式
	 * 
	 * @see 将行样式对应单元格进行复制
	 * @author lfh
	 * @param rowk 原行
	 * @param rowi 目标行
	 * @param s 开始列
	 * @param e 结束列
	 */
	public void copyRowStyle(int rowk, int rowi, int s, int e) {
		XSSFRow _rowk = this.entryRow(rowk);
		XSSFRow _rowi = this.entryRow(rowi);
		this.copyRowStyle(_rowk, _rowi, s, e);
	}

	/**
	 * @author LFH
	 * @date 2017年5月17日 下午6:04:58
	 * @see 设置行样式
	 * @param rowi 行标
	 * @param s 列开始
	 * @param e 列结束
	 * @param style
	 */
	public void setRowStyle(int rowi, int s, int e, XSSFCellStyle style) {
		if (this.tSheet == null) {
			System.err.println("请先创建Sheet!");
			return;
		}
		for (; s <= e; s++) {
			XSSFCell cell = entryCell(entryRow(rowi), s);
			cell.setCellStyle(style);
		}
	}

	/**
	 * @author LFH
	 * @date 2017年5月17日 下午6:04:58
	 * @see 设置当前sheet中行样式
	 * @param rowi 行标
	 * @param s 列开始
	 * @param e 列结束
	 * @param style
	 * @param times 次数
	 */
	public void setRowStyle(int rowi, int s, int e, XSSFCellStyle style, int times) {
		if (this.tSheet == null) {
			System.err.println("请先创建Sheet!");
			return;
		}
		for (int i = 0; i < times; i++) {
			setRowStyle(rowi + i, s, e, style);
		}
	}

	/**
	 * @author LFH
	 * @date 2017年5月17日 下午6:25:43
	 * @see 创建CellStyle,以供额外的设置样式
	 * @return
	 */
	public XSSFCellStyle createStyle() {
		if (this.tWorkbook == null) {
			System.err.println("请先创建Sheet!");
			return null;
		}
		return this.xssfCellStyle;
	}

	/**
	 * 创建数据格式
	 * 
	 * @author LFH
	 * @date 2018年3月2日 上午10:31:37
	 * @return
	 */
	public XSSFDataFormat createDataFormat() {
		if (this.tWorkbook == null) {
			System.err.println("请先创建Sheet!");
			return null;
		}
		return this.tWorkbook.createDataFormat();
	}

	/**
	 * @author LFH
	 * @date 2017年5月17日 下午6:38:33
	 * @see 获取CellStyle,以供额外的附加样式
	 * @param rowi 行标
	 * @param celli 列标
	 * @return
	 */
	public XSSFCellStyle getStyle(int rowi, int celli) {
		if (this.tSheet == null) {
			System.err.println("请先创建Sheet!");
			return null;
		}
		XSSFCell cell = this.entryCell(this.entryRow(rowi), celli);
		XSSFCellStyle style = cell.getCellStyle();
		return style;
	}

	/**
	 * @author LFH
	 * @date 2017年5月17日 下午1:52:49
	 * @see 复制单元格
	 * @param fromRow 源行
	 * @param fromCell 源列
	 * @param toRow 目的行
	 * @param toCell 目的列
	 * @param copyValueFlag 是否包含内容
	 */
	public void copyCell(int fromRow, int fromCell, int toRow, int toCell, boolean copyValueFlag) {
		if (this.tSheet == null) {
			System.err.println("请先创建Sheet!");
			return;
		}
		XSSFCell s = entryCell(entryRow(fromRow), fromCell);
		XSSFCell t = entryCell(entryRow(toRow), toCell);
		copyCell(s, t, copyValueFlag);
	}

	/**
	 * @author LFH
	 * @see 复制行
	 * @param fromRow
	 * @param toRow
	 * @param copyValueFlag 是否含内容复制
	 */
	public void copyRow(int fromRow, int toRow, boolean copyValueFlag) {
		if (this.tSheet == null) {
			System.err.println("请先创建Sheet!");
			return;
		}
		XSSFRow f = entryRow(fromRow);
		XSSFRow t = entryRow(toRow);
		copyRow(f, t, copyValueFlag);
	}

	/**
	 * @author LFH
	 * @date 2017年5月17日 下午2:01:07
	 * @see 复制行(指定范围列)
	 * @param fromRow 源行
	 * @param toRow 目的行
	 * @param fromCell 复制列的起始
	 * @param toCell 复制列的结束
	 * @param copyValueFlag
	 */
	public void copyRow(int fromRow, int toRow, int fromCell, int toCell, boolean copyValueFlag) {
		for (int i = fromCell; i <= toCell; i++) {
			copyCell(fromRow, i, toRow, i, copyValueFlag);
		}
	}

	/**
	 * @author LFH
	 * @date 2017年5月17日 下午4:58:56
	 * @see 清空行的值
	 * @param fromCell 起始列
	 * @param endCell 结束列
	 * @param rowk 待操作行
	 */
	public void clearRow(int fromCell, int endCell, int... rowk) {
		if (this.tSheet == null) {
			System.err.println("请先创建Sheet!");
			return;
		}
		for (int i : rowk) {
			XSSFRow rowi = entryRow(i);
			for (; fromCell <= endCell; fromCell++) {
				XSSFCell c = entryCell(rowi, fromCell);
				c.setCellType(XSSFCell.CELL_TYPE_BLANK);
			}
		}
	}

	/**
	 * @author LFH
	 * @date 2017年5月17日 下午2:01:07
	 * @see 复制行(指定范围列)[可复制多次]
	 * @param fromRow 源行
	 * @param toRow1 目的行1
	 * @param times 复制次数(从目的行开始计算复制次数)
	 * @param fromCell 复制列的起始
	 * @param toCell 复制列的结束
	 * @param copyValueFlag 是否包含内容
	 * @param clear 是否清空内容(可选)
	 */
	public void copyRow(int fromRow, int toRow, int times, int fromCell, int toCell, boolean copyValueFlag,
			Boolean... clear) {
		for (int it = 0; it <= times; it++) {
			copyRow(fromRow, toRow++, fromCell, toCell, copyValueFlag);
			if (clear != null && clear.length > 0 && clear[0] == true) {
				clearRow(fromCell, toCell, toRow - 1);
			}
		}
	}

	/**
	 * 移动行(携带样式和内容)
	 * 
	 * @author LFH
	 * @date 2018年3月1日 上午10:07:45
	 * @param startRow 移动区域起始行
	 * @param endRow 移动区域截止行
	 * @param n 移动几行
	 */
	public void shiftRows(int startRow, int endRow, int n) {
		this.tSheet.shiftRows(startRow, endRow, n, true, true);
	}
	
	/**
	 * 创建sheet,返回序号
	 * 
	 * @author LFH
	 * @date 2018年3月23日 下午5:45:06
	 * @return 序号
	 */
	public int createSheet() {
		XSSFSheet sheet = this.tWorkbook.createSheet();
		return this.tWorkbook.getSheetIndex(sheet);
	}

	/**
	 * 创建sheet ,并设置名称,返回序号.
	 * 
	 * @author LFH
	 * @date 2018年3月23日 下午5:44:46
	 * @param name
	 * @return 序号
	 */
	public int createSheet(String name) {
		XSSFSheet sheet = this.tWorkbook.createSheet(name);
		return this.tWorkbook.getSheetIndex(sheet);
	}

	/**
	 * 根据序号移除sheet
	 * 
	 * @author LFH
	 * @date 2018年3月23日 下午5:41:55
	 * @param at
	 * @return true/false
	 */
	public boolean removeSheet(int at) {
		boolean b = false;
		if (this.tWorkbook.getSheetAt(at) != null) {
			this.tWorkbook.removeName(at);
			b = true;
		}
		return b;
	}

	/**
	 * 根据名称移除sheet
	 * 
	 * @author LFH
	 * @date 2018年3月23日 下午5:42:07
	 * @param name
	 * @return true/false
	 */
	public boolean removeSheet(String name) {
		boolean b = false;
		if (this.tWorkbook.getSheetIndex(name) > -1) {
			this.tWorkbook.removeName(name);
			b = true;
		}
		return b;
	}
/************************************************************************************************/
	/** 下方私有方法请勿设置公开! */

	/**
	 * 复制行样式
	 * 
	 * @see 将行样式对应单元格进行复制
	 * @author lfh
	 * @param rowk 原行
	 * @param rowi 目标行
	 * @param s 开始列
	 * @param e 结束列
	 */
	private void copyRowStyle(XSSFRow rowk, XSSFRow rowi, int s, int e) {
		if (rowk == null) {
			return;
		}
		XSSFCellStyle style = this.createStyle();
		XSSFCellStyle rowStyle = rowk.getRowStyle();
		if (rowStyle != null) {
			rowi.setHeight(rowk.getHeight());
			rowi.setRowStyle(rowStyle);				
		}
		XSSFCell cell = null;
		int i = s;
		for (; i <= e; i++) {
			style = entryCell(rowk, i).getCellStyle();
			cell = entryCell(rowi, i);
			cell.setCellStyle(style);
		}
		return;
	}

	/**
	 * 获得有效单元格
	 * 
	 * @author LFH
	 * @param
	 * @param
	 * @return XSSFcell
	 */
	private XSSFCell entryCell(XSSFRow rowi, int index) {
		XSSFCell cell = rowi.getCell(index);
		if (cell == null) {
			cell = rowi.createCell(index);
		}
		return cell;
	}

	/**
	 * 填值前预判是否为空
	 * 
	 * @author LFH
	 * @date 2018年4月14日 下午10:33:05
	 * @param value
	 * @return
	 */
	private static boolean isNull(Object value) {
		return value == null || "null".equals(value) || "".equals(value);
	}
	/**
	 * 获取有效单元格并设置值(含样式)
	 * 
	 * @author LFH
	 * @date 2017年5月12日 下午2:47:04
	 * @param rowi 行
	 * @param index 列标
	 * @param value 值
	 * @param style 样式
	 * @return
	 */
	private void entryCell(XSSFRow rowi, int index, Object value, XSSFCellStyle style) {
		XSSFCell cell = entryCell(rowi, index);
		cell.setCellStyle(style);
		String v = "";
		if (isNull(value)) {
			v = "";
		} else {
			v = value.toString();
		}
		cell.setCellValue(v);
	}


	/**
	 * 获取有效单元格并设置值(数值类型);
	 * 
	 * @see 非数值则为空
	 * @author LFH
	 * @param rowi 行
	 * @param index 列标
	 * @param value 值
	 */
	private void numCell(XSSFRow rowi, int index, Object value) {
		XSSFCell cell = entryCell(rowi, index);
		String v = "";
		if (isNull(value)) {
			v = "0";
		} else {
			v = value.toString();
		}
		if (v.indexOf(".") != -1) {
			cell.setCellValue(Double.parseDouble(v));
		} else if (!"0".equals(v)) {
			cell.setCellValue(Integer.parseInt(v));
		} else {
			cell.setCellValue("");
		}
	}

	/**
	 * 获取有效单元格并设置值(数值类型);
	 * 
	 * @see 非数值则为0
	 * @author LFH
	 * @param rowi 行
	 * @param index 列标
	 * @param value 值
	 */
	private void numCellZ(XSSFRow rowi, int index, Object value) {
		XSSFCell cell = entryCell(rowi, index);
		String v = "";
		if (isNull(value)) {
			v = "0";
		} else {
			v = value.toString();
		}
		if (v.indexOf(".") != -1) {
			cell.setCellValue(Double.parseDouble(v));
		} else {
			cell.setCellValue(Integer.parseInt(v));
		}
	}

	/**
	 * 非数值或零值则为空.
	 * 
	 * @author LFH
	 * @date 2017年12月8日 下午10:30:34
	 * @param rowi
	 * @param index
	 * @param value
	 */
	private void numCellT(XSSFRow rowi, int index, Object value) {
		XSSFCell cell = entryCell(rowi, index);
		String v = "";
		if (isNull(value)) {
			v = "0";
		} else {
			v = value.toString();
		}
		if (v.indexOf(".") != -1) {
			cell.setCellValue(Double.parseDouble(v));
		} else if ("0".equals(v)) {
			cell.setCellValue("");
		} else {
			cell.setCellValue(Integer.parseInt(v));
		}
	}

	/**
	 * 获取有效单元格并设置值
	 * 文本值
	 * 
	 * @author LFH
	 * @param rowi 行
	 * @param index 列标
	 * @param value 值
	 * @return
	 */
	private void entryCell(XSSFRow rowi, int index, Object value) {
		XSSFCell cell = entryCell(rowi, index);
		String v = "";
		if (isNull(value)) {
			v = "";
		} else {
			v = value.toString();
		}
		cell.setCellValue(v);
	}

	private void entryCellRich(XSSFRow rowi, int index, Object value) {
		XSSFCell cell = entryCell(rowi, index);
		String v = "";
		if (isNull(value)) {
			v = "";
		} else {
			v = value.toString();
		}
		cell.setCellValue(new XSSFRichTextString(v));
	}

	/**
	 * 日期单元格
	 * 
	 * @author LFH
	 * @param
	 * @param
	 * @param
	 */
	private void dateCell(XSSFRow rowi, int index, Date value) {
		if (!(value instanceof Date)) {
			System.err.println("请传入日期类型数据!");
			return;
		}
		XSSFCell cell = entryCell(rowi, index);
		Date v = value;
		if (isNull(value)) {
			v = null;
		} else {
			v = value;
		}
		cell.setCellValue(v);
	}

	/**
	 * 公式单元格
	 * 
	 * @author LFH
	 * @param
	 * @param
	 * @param
	 */
	private void formulaCell(XSSFRow rowi, int index, Object value) {
		XSSFCell cell = entryCell(rowi, index);
		String v = "";
		if (isNull(value)) {
			v = "";
			cell.setCellValue(v);
		} else {
			try {
				v = value.toString();
				cell.setCellFormula(v);
			} catch (Exception e) {
				// TODO: handle exception
				System.err.println(e);
				cell.setCellValue(v);
			}

		}
	}

	/**
	 * 向浏览器发送文件流,以供下载
	 * 
	 * @author LFH
	 * @date 2017年5月12日 下午2:47:04
	 * @param response
	 * @param out 浏览器输出流
	 * @param fileName 文件名(不用传后缀)
	 * @param workbook excel文档
	 */
	private void outFile(HttpServletResponse response, String fname, XSSFWorkbook wb) throws Exception {
		/**********************/
		OutputStream out = null;
		out = response.getOutputStream();
		fname = fname == null || fname.trim().length() <= 0 ? "_" : fname;
		fname = new String(fname.getBytes("gbk"), "iso-8859-1");
		response.reset();
		response.setContentType("application/vnd.ms-excel;charset=utf-8");
		// 靠这一行向外导出文件,("文件名以双引号包住,是为了避免在火狐等浏览器导出文件时文件名中有空格出现问题的情况 !")
		response.setHeader("Content-disposition", "attachment; filename=\"" + fname + ".xls" + "\"");
		wb.write(out);
		if (out != null) {
			out.close();
		}
	}

	/**
	 * @author LFH
	 * @date 2017年5月12日 下午2:47:04
	 * @see 获得文档流
	 * @param request
	 * @param path
	 * @return fis 文件流
	 * @throws Exception
	 */
	private static FileInputStream fileInput(String rootPath, String path) throws Exception {
		String filePath = null;
		File file = null;
		FileInputStream fis;
		filePath = rootPath + path;
		file = new File(filePath);
		fis = new FileInputStream(file);
		return fis;
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

	/**
	 * 私有方法用于批量填值用.
	 * 
	 * @author LFH
	 * @param rowi 行号
	 * @param c 列号
	 * @param v 值
	 * @param type 类型
	 */
	private void entryCellInType(XSSFRow rowi, int c, Object v, Type type) {
		type = type == null ? Type.NORMAL : type;
		switch (type) {
		case NORMAL:
			entryCell(rowi, c, v);/** 常规 */
			break;
		case NUM:
			numCell(rowi, c, v);/** 非数值则空 */
			break;
		case NUMZ:
			numCellZ(rowi, c, v);/** 非数值则0 */
			break;
		case NUMT:
			numCellT(rowi, c, v);/** 非数值或0值则空 */
			break;
		case DATE:
			dateCell(rowi, c, (Date) v);/** 日期 */
			break;
		case FORMULA:
			formulaCell(rowi, c, v);/** 公式 */
			break;
		case RICH:
			entryCellRich(rowi, c, v);/** 富文本 */
			break;
		default:
			entryCell(rowi, c, v);
			break;
		}
	}

	/**
	 * 复制单元格
	 * 
	 * @author LFH
	 * @param srcCell
	 * @param distCell
	 * @param copyValueFlag
	 *            是否含内容
	 */
	private void copyCell(XSSFCell srcCell, XSSFCell distCell, boolean copyValueFlag) {
		// distCell.setEncoding(srcCell.getEncoding());
		// 目标单元格样式设置
		distCell.setCellStyle(srcCell.getCellStyle());
		//
		if (srcCell.getCellComment() != null) {
			distCell.setCellComment(srcCell.getCellComment());
		}
		// 单元格复制
		int srcCellType = srcCell.getCellType();
		distCell.setCellType(srcCellType);
		if (copyValueFlag) {
			if (srcCellType == XSSFCell.CELL_TYPE_NUMERIC) {
				if (HSSFDateUtil.isCellDateFormatted(srcCell)) {
					distCell.setCellValue(srcCell.getDateCellValue());
				} else {
					distCell.setCellValue(srcCell.getNumericCellValue());
				}
			} else if (srcCellType == XSSFCell.CELL_TYPE_STRING) {
				distCell.setCellValue(srcCell.getRichStringCellValue());
			} else if (srcCellType == XSSFCell.CELL_TYPE_BLANK) {
				//
			} else if (srcCellType == XSSFCell.CELL_TYPE_BOOLEAN) {
				distCell.setCellValue(srcCell.getBooleanCellValue());
			} else if (srcCellType == XSSFCell.CELL_TYPE_ERROR) {
				distCell.setCellErrorValue(srcCell.getErrorCellValue());
			} else if (srcCellType == XSSFCell.CELL_TYPE_FORMULA) {
				distCell.setCellFormula(srcCell.getCellFormula());
			} else { //
			}
		}
	}

	/**
	 * @author LFH
	 *         复制行
	 * @param fromRow
	 * @param toRow
	 * @param copyValueFlag 是否含内容复制
	 */
	private void copyRow(XSSFRow fromRow, XSSFRow toRow, boolean copyValueFlag) {
		for (Iterator<Cell> cellIt = fromRow.cellIterator(); cellIt.hasNext();) {
			XSSFCell tmpCell = (XSSFCell) cellIt.next();
			XSSFCell newCell = toRow.createCell(tmpCell.getColumnIndex());
			copyCell(tmpCell, newCell, copyValueFlag);
		}
	}

	/**
	 * @author LFH
	 * @date 2017年5月12日
	 * @see 自定义的填充数据类型
	 */
	public enum Type {
		NUM, NUMZ, NUMT, DATE, FORMULA, NORMAL, RICH

	}

	/**
	 * @author LFH
	 * @date 2017年5月12日 下午5:52:38
	 * @see 主要为方便批量设置数据类型
	 * @param map 传入Map
	 * @param t 数据类型
	 * @param i 对应列号
	 * @return
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
	 * @author LFH
	 * @date 2017年12月8日 下午10:19:40
	 * @param t
	 * @param i
	 * @return
	 */
	public Map<Integer, Type> mapType(Type t, Integer... i) {
		Map<Integer, Type> types = new HashMap<>();
		return mapType(types, t, i);
	}

	/**
	 * 获取类型Map集合.(静态方法)
	 * 
	 * @author LFH
	 * @date 2018年3月20日 下午4:17:10
	 * @param type
	 * @param t
	 * @param i
	 * @return
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
	 * @author LFH
	 * @date 2018年3月20日 下午4:16:52
	 * @param t
	 * @param i
	 * @return
	 */
	public static Map<Integer, Type> mapTypes(Type t, Integer... i) {
		Map<Integer, Type> types = new HashMap<>();
		return mapTypes(types, t, i);
	}
/**
	 * 快速创建位置对象
	 * 
	 * @author LFH
	 * @date 2018年3月7日 下午9:24:16
	 * @return
	 */
	public CellPosition cps(String key, int row, int cell) {
		return new CellPosition(key, row, cell);
	}

	/**
	 * 快速创建位置对象
	 * 
	 * @author LFH
	 * @date 2018年3月11日 下午3:28:05
	 * @return
	 */
	public CellPosition cps(String key, int row, int cell, Type type) {
		return new CellPosition(key, row, cell, type);
	}
	
	/**
	 * 快速创建位置对象(附加值)
	 * 
	 * @author LFH
	 * @date 2018年3月11日 下午3:28:05
	 * @return
	 */
	public CellPosition cps(int row, int cell, Object value) {
		return new CellPosition(row, cell, value);
	}

	/**
	 * 快速创建位置对象(附加值)
	 * 
	 * @author LFH
	 * @date 2018年3月11日 下午3:28:05
	 * @return
	 */
	public CellPosition cps(int row, int cell, Object value, Type type) {
		return new CellPosition(row, cell, value, type);
	}

	/**
	 * 快速创建批量位置集合
	 * 
	 * @see -- 对象参数p 由
	 * @see #cps(String, int, int)
	 * @see #cps(String, int, int, Type)
	 * @see -- 方法创建
	 * @author LFH
	 * @date 2018年3月7日 下午9:28:11
	 * @param p
	 * @return
	 */
	public List<CellPosition> cpsBatch(CellPosition... p) {
		List<CellPosition> pos = new ArrayList<>(Arrays.asList(p));
		return pos;
	}

	/**
	 * 填充sheet时用到的位置对象.
	 * 
	 * @author LFH
	 * @date 2018年3月7日
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
	 * @author LFH
	 *         在表格中插入图片
	 * @see 目前支持jpg/jpeg/gif/png/bmp格式图片
	 * @param request
	 * @param XSSFWorkbook
	 * @param sheet1
	 * @param fPath 图片地址
	 * @param sr 图片开始行
	 * @param er 图片结束行
	 * @param sc 图片开始列
	 * @param ec 图片结束列
	 */
	public void insertImg(ImgOption option) throws Exception {
		String type = "";
		String[] types = { "jpg", "png", "gif", "bmp", "jpeg" };
		boolean check = false;
		String fPath = option.getfPath();
		if (fPath != null && fPath.indexOf(".") != -1) {
			type = fPath.substring(fPath.lastIndexOf(".") + 1, fPath.length());
			type = type.trim();
			for (String t : types) {
				if (t.equalsIgnoreCase(type)) {
					check = true;
				}
			}
			if (!check) {
				throw new Exception("Image's type is not defined or the type can't be read!\t" + type);
			}
		} else {
			throw new Exception("Image File Not Found Exception!\t" + type);
		}
		// 新建图片缓存区
		BufferedImage bufferImg = null;
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		bufferImg = ImageIO.read(new File(this.rootPath + fPath));
		ImageIO.write(bufferImg, type, byteArrayOut);
		XSSFDrawing patriarch = this.tSheet.createDrawingPatriarch();
		// anchor对象创建
		XSSFClientAnchor anchor = new XSSFClientAnchor(option.getDx1(), option.getDy1(), option.getDx2(),
				option.getDy2(), option.getCol1(), option.getRow1(), option.getCol2(), option.getRow2());
		anchor.setAnchorType(AnchorType.MOVE_AND_RESIZE);
		// anchor.setAnchorType(3); //3.16POI有更改.
		// 插入图片
		patriarch.createPicture(anchor,
				this.tWorkbook.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_PNG));
	}

	/** 图片属性 */
	public ImgOption initImgOption(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2,
			String path) {
		return new ImgOption(dx1, dy1, dx2, dy2, (short) col1, row1, (short) col2, row2, path);
	}

	public ImgOption initImgOption(int dx2, int dy2, int col1, int row1, int col2, int row2, String path) {
		return new ImgOption(dx2, dy2, (short) col1, row1, (short) col2, row2, path);
	}

	public ImgOption initImgOption(int col1, int row1, int col2, int row2, String path) {
		return new ImgOption((short) col1, row1, (short) col2, row2, path);
	}

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

		public String getfPath() {
			return fPath;
		}

		public int getDx1() {
			return dx1;
		}

		public int getDy1() {
			return dy1;
		}

		public int getDx2() {
			return dx2;
		}

		public int getDy2() {
			return dy2;
		}

		public short getCol1() {
			return col1;
		}

		public int getRow1() {
			return row1;
		}

		public short getCol2() {
			return col2;
		}

		public int getRow2() {
			return row2;
		}

		private ImgOption(int dx1, int dy1, int dx2, int dy2, short col1, int row1, short col2, int row2, String path) {
			this.dx1 = dx1;
			this.dy1 = dy1 > 255 ? 255 : dy1;
			this.dx2 = dx2;
			this.dy2 = dy2 > 255 ? 255 : dy2;
			this.col1 = col1;
			this.row1 = row1;
			this.col2 = col2;
			this.row2 = row2;
			this.fPath = path;
		}

		private ImgOption(int dx2, int dy2, short col1, int row1, short col2, int row2, String path) {
			this.dx2 = dx2;
			this.dy2 = dy2;
			this.col1 = col1;
			this.row1 = row1;
			this.col2 = col2;
			this.row2 = row2;
			this.fPath = path;
		}

		private ImgOption(short col1, int row1, short col2, int row2, String path) {
			this.col1 = col1;
			this.row1 = row1;
			this.col2 = col2;
			this.row2 = row2;
			this.fPath = path;
		}

	}

	/** 单元格简单设置操作工具类 */
	/**
	 * 创建一个配置样式的对象.
	 * 注意:此设置应尽量在批量填值完成后操作.
	 * 
	 * @example
	 * 			XSSFCellStyle style = exp.getStyle(6, 1);<br/>
	 *          exp.entryCell(6, 0).setCellStyle( exp.createExpStyle(style)<br/>
	 *          .setFontColor(XSSFColor.RED.index)<br/>
	 *          .setFgColor(XSSFColor.YELLOW.index)<br/>
	 *          .setFontSize(20).setAlign("right", "bottom").finish());
	 * @author LFH
	 * @date 2018年4月14日 下午9:09:11
	 * @param expStyle
	 * @return
	 */
	public ExpStyle createExpStyle(XSSFCellStyle expStyle) {
		return new ExpStyle(this.tWorkbook, expStyle);
	}

	/**
	 * @see --示例参看:
	 * @see #createExpStyle(XSSFCellStyle)
	 * 
	 * @author LFH
	 * @date 2018年4月14日 下午10:21:58
	 * @return
	 */
	public ExpStyle createExpStyle() {
		return new ExpStyle(this.tWorkbook);
	}

	/**
	 * 对齐方式变量
	 * 
	 * @author LFH
	 * @date 2018年4月14日
	 */
	private enum Align {
		center(XSSFCellStyle.ALIGN_CENTER), left(XSSFCellStyle.ALIGN_LEFT), right(XSSFCellStyle.ALIGN_RIGHT), middle(
				XSSFCellStyle.ALIGN_CENTER), top(XSSFCellStyle.VERTICAL_TOP), bottom(XSSFCellStyle.VERTICAL_BOTTOM);
		private short index;

		private Align(short index) {
			this.index = index;
		}

	}

	/**
	 * 样式配置类
	 * 
	 * @author LFH
	 * @date 2018年4月14日
	 */
	public class ExpStyle {
		private XSSFCellStyle expStyle;
		private XSSFFont expFont;
		private boolean hasFont = false;

		private ExpStyle(XSSFWorkbook work) {
			this.expStyle = work.createCellStyle();
			this.expFont = work.createFont();
		}

		private ExpStyle(XSSFWorkbook work, XSSFCellStyle expStyle) {
			this.expStyle = work.getNumCellStyles() < 3500 ? work.createCellStyle()
					: work.getCellStyleAt((short) (work.getNumCellStyles() - 1));
			Class<?> type = expStyle.getClass();
			try {
				BeanInfo beanInfo = Introspector.getBeanInfo(type);
				PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();// 获取属性数组
				for (PropertyDescriptor pd : propertyDescriptors) {
					String propertyName = pd.getName();// 获取属性名
					if (!"class".equalsIgnoreCase(propertyName)) {
						Method get = pd.getReadMethod();// 得到读属性方法(get...())
						Method set = pd.getWriteMethod();
						if (get != null) {
							Object value = get.invoke(expStyle);// 获取属性值
							if (value != null && set != null) {
								set.invoke(this.expStyle, value);
							}
						}
					}
				}
			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
			}
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
		 * @author LFH
		 * @date 2018年4月14日 下午8:48:22
		 * @param color {@link XSSFColor} 的<b>颜色属性</b> 的 <b>index</b> 属性
		 * @return
		 */
		public ExpStyle setFontColor(short color) {
			this.expFont.setColor(color);
			addFont();
			return this;
		}

		/**
		 * 添加字体大小
		 * 
		 * @author LFH
		 * @date 2018年4月14日 下午8:56:36
		 * @param size
		 * @return
		 */
		public ExpStyle setFontSize(int size) {
			this.expFont.setFontHeightInPoints((short) size);
			addFont();
			return this;
		}

		/**
		 * 设置单元格对齐
		 * 
		 * @author LFH
		 * @date 2018年4月14日 下午8:59:36
		 * @param tAlign 水平
		 * @param vAlign 垂直
		 * @return
		 */
		public ExpStyle setAlign(String tAlign, String vAlign) {
			try {
				Align talign = Align.valueOf(tAlign);
				this.expStyle.setAlignment(talign.index);
			} catch (Exception e) {
				// TODO: handle exception
			}
			try {
				Align valign = Align.valueOf(vAlign);
				this.expStyle.setVerticalAlignment(valign.index);
			} catch (Exception e) {
				// TODO: handle exception
			}
			return this;
		}

		/**
		 * 设置背景颜色
		 * 
		 * @author LFH
		 * @date 2018年4月14日 下午9:11:59
		 * @param color {@link XSSFColor} 的<b>颜色属性</b> 的 <b>index</b> 属性
		 * @return
		 */
		public ExpStyle setBgColor(short color) {
			this.expStyle.setFillBackgroundColor(color);
			this.expStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
			return this;
		}

		/**
		 * 设置前景颜色
		 * 
		 * @author LFH
		 * @date 2018年4月14日 下午9:11:59
		 * @param color {@link XSSFColor} 的<b>颜色属性</b> 的 <b>index</b> 属性
		 * @return
		 */
		public ExpStyle setFgColor(short color) {
			this.expStyle.setFillForegroundColor(color);
			this.expStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
			return this;
		}

		/**
		 * 完成设置,返回样式.
		 * 
		 * @author LFH
		 * @date 2018年4月14日 下午9:59:53
		 * @return
		 */
		public XSSFCellStyle finish() {
			return this.expStyle;
		}
	}
}