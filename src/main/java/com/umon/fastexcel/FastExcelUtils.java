package com.umon.fastexcel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.umon.fastexcel.annotation.ExcelCell;
import com.umon.fastexcel.annotation.ExcelSheet;


/**
 * 
 * @className:FastExcelUtils
 * @description:
 * <p>
 * 快速简单操作Excel的工具
 * </p>
 * @author qinzy
 * @datetime:2017年11月6日
 *
 */
public class FastExcelUtils {

	private static final Logger logger = LoggerFactory.getLogger(FastExcelUtils.class);

	/**
	 * 时日类型的数据默认格式化方式
	 */
	private static DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

	public static final String XLS = "xls";
	public static final String XLSX = "xlsx";

	/**
	 * 从文件路径导入Excel文件，并封装成对象
	 *
	 * @param sheetClass
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static <T> List<T> importExcel(Class<T> sheetClass, String filePath) throws IOException {
		File excelFile = new File(filePath);
		if (!excelFile.exists()) {
			logger.warn("文件:{} 不存在！创建此文件！" + filePath);
			if (!excelFile.createNewFile()) {
				throw new IOException("文件创建失败");
			}
		}
		List<T> dataList = importExcel(sheetClass, excelFile);
		return dataList;
	}

	/**
	 * 导入Excel文件，并封装成对象
	 *
	 * @param sheetClass
	 * @param excelFile
	 * @return
	 */
	public static <T> List<T> importExcel(Class<T> sheetClass, File excelFile) {
		Workbook workbook = null;
		try {
			if (excelFile == null || !excelFile.exists()) {
				logger.warn("文件:{} 不存在！");
				throw new IOException("文件:{} 不存在！文件创建失败");
			}
			workbook = WorkbookFactory.create(excelFile);
			List<T> dataList = importExcel(sheetClass, workbook);
			return dataList;
		} catch (IOException e) {
			logger.error(e.getMessage(), e);
			throw new RuntimeException(e);
		} catch (InvalidFormatException e) {
			logger.error(e.getMessage(), e);
			throw new RuntimeException(e);
		} finally {
			try {
				if (workbook != null) {
					workbook.close();
				}
			} catch (IOException e) {
				logger.error(e.getMessage(), e);
				throw new RuntimeException(e);
			}
		}
	}

	/**
	 * 导入Excel数据流，并封装成对象
	 *
	 * @param sheetClass
	 * @param inputStream
	 * @return
	 */
	public static <T> List<T> importExcel(Class<T> sheetClass, InputStream inputStream) {
		Workbook workbook = null;
		try {
			workbook = WorkbookFactory.create(inputStream);
			List<T> dataList = importExcel(sheetClass, workbook);
			return dataList;
		} catch (IOException e) {
			logger.error(e.getMessage(), e);
			throw new RuntimeException(e);
		} catch (InvalidFormatException e) {
			logger.error(e.getMessage(), e);
			throw new RuntimeException(e);
		} finally {
			try {
				if (workbook != null) {
					workbook.close();
				}
			} catch (IOException e) {
				logger.error(e.getMessage(), e);
				throw new RuntimeException(e);
			}
		}
	}

	/**
	 * 从Workbook导入Excel文件，并封装成对象
	 *
	 * @param sheetClass
	 * @param workbook
	 * @return
	 */
	public static <T> List<T> importExcel(Class<T> sheetClass, Workbook workbook) {
		List<T> datas = null;
		String sheetName = null;
		if (sheetClass.isAnnotationPresent(ExcelSheet.class)) {
			ExcelSheet excelSheet = sheetClass.getAnnotation(ExcelSheet.class);
			if (sheetName == null || "".equals(sheetName.trim())) {
				if (excelSheet.name() != null && !"".equals(excelSheet.name().trim())) {
					sheetName = excelSheet.name().trim();
				} else if (excelSheet.value() != null && !"".equals(excelSheet.value().trim())) {
					sheetName = excelSheet.value().trim();
				}
			}
		}
		if (sheetName != null) {
			Sheet sheet = workbook.getSheet(sheetName);
			if (null != sheet) {
				datas = readExcel(sheetClass, sheet);
			} else {
				throw new RuntimeException("sheetName:" + sheetName + " is not exist");
			}
		} else {
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				Sheet sheet = workbook.getSheetAt(i);
				if (datas == null) {
					datas = readExcel(sheetClass, sheet);
				} else {
					datas.addAll(readExcel(sheetClass, sheet));
				}
			}
		}

		return datas;
	}

	private static <T> List<T> readExcel(Class<T> sheetClass, Sheet sheet) {
		List<T> datas = null;
		if (null != sheet) {
			datas = new ArrayList<T>();
			int startRow = 0;
			// sheet field
			Map<String, Field> fieldMap = new HashMap<String, Field>();
			Map<String, String> titleMap = new HashMap<String, String>();
			Field[] fields = sheetClass.getDeclaredFields();
			for (Field field : fields) {
				if (Modifier.isStatic(field.getModifiers())) {
					continue;
				}
				if (field.isAnnotationPresent(ExcelCell.class)) {
					ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
					String cellName = "";
					if (excelCell.name() != null && !"".equals(excelCell.name().trim())) {
						cellName = excelCell.name();
					} else if (excelCell.value() != null && !"".equals(excelCell.value().trim())) {
						cellName = excelCell.value();
					} else {
						cellName = field.getName();
					}
					fieldMap.put(cellName, field);
				}
			}

			if (fieldMap.isEmpty()) {
				throw new RuntimeException("---------> Fast-excel error, data field has not been annotation.");
			}

			Row row = sheet.getRow(startRow);
			for (Cell cell : row) {
				// 查找有公式的单元格
				CellReference cellRef = new CellReference(cell);
				titleMap.put(cellRef.getCellRefParts()[2], cell.getRichStringCellValue().getString());
			}

			for (int i = startRow + 1; i <= sheet.getLastRowNum(); i++) {
				try {
					T t = sheetClass.newInstance();
					Row rowX = sheet.getRow(i);
					for (Cell cell : rowX) {
						CellReference cellRef = new CellReference(cell);
						String cellTag = cellRef.getCellRefParts()[2];
						String name = titleMap.get(cellTag);
						Field field = fieldMap.get(name);
						if (null != field) {
							field.setAccessible(true);
							setFieldValue(cell, t, field);
						}
					}
					datas.add(t);
					logger.debug(t.toString());
				} catch (InstantiationException | IllegalAccessException | SecurityException
						| IllegalArgumentException | ParseException e) {
					e.printStackTrace();
					logger.error(e.getMessage(), e);
					throw new RuntimeException(e);
				}
			}
		} else {
			throw new RuntimeException("sheet is not exist, it is null");
		}
		return datas;
	}

	private static void setFieldValue(Cell cell, Object obj, Field field)
			throws IllegalAccessException, ParseException {
		Class<?> fieldType = field.getType();
		switch (cell.getCellTypeEnum()) {
		case BLANK:
			break;
		case BOOLEAN:
			field.setBoolean(obj, cell.getBooleanCellValue());
			break;
		case ERROR:
			field.setByte(obj, cell.getErrorCellValue());
			break;
		case FORMULA:
			field.set(obj, cell.getCellFormula());
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				if (field.getType().getName().equals(Date.class.getName())) {
					field.set(obj, cell.getDateCellValue());
				} else {
					field.set(obj, dateFormat.format(cell.getDateCellValue()));
				}
			} else {
				if (fieldType.isAssignableFrom(Integer.class) || Integer.class.equals(fieldType)
						|| Integer.TYPE.equals(fieldType)) {
					field.setInt(obj, (int) cell.getNumericCellValue());
				} else if (fieldType.isAssignableFrom(Short.class) || Short.class.equals(fieldType)
						|| Short.TYPE.equals(fieldType)) {
					field.setShort(obj, (short) cell.getNumericCellValue());
				} else if (fieldType.isAssignableFrom(Float.class) || Float.class.equals(fieldType)
						|| Float.TYPE.equals(fieldType)) {
					field.setFloat(obj, (float) cell.getNumericCellValue());
				} else if (fieldType.isAssignableFrom(Byte.class) || Byte.class.equals(fieldType)
						|| Byte.TYPE.equals(fieldType)) {
					field.setByte(obj, (byte) cell.getNumericCellValue());
				} else if (fieldType.isAssignableFrom(Double.class) || Double.class.equals(fieldType)
						|| Double.TYPE.equals(fieldType)) {
					field.setDouble(obj, cell.getNumericCellValue());
				} else if (fieldType.isAssignableFrom(String.class)) {
					String s = String.valueOf(cell.getNumericCellValue());
					if (s.contains("E")) {
						s = s.trim();
						BigDecimal bigDecimal = new BigDecimal(s);
						s = bigDecimal.toPlainString();
					}
					// 防止整数判定为浮点数
					if (s.endsWith(".0")) {
						s = s.substring(0, s.indexOf(".0"));
					}
					field.set(obj, s);
				} else {
					field.set(obj, cell.getNumericCellValue());
				}
			}
			break;
		case STRING:
			if (fieldType.getName().equals(Date.class.getName())) {
				field.set(obj, dateFormat.parse(cell.getRichStringCellValue().getString()));
			} else {
				field.set(obj, cell.getRichStringCellValue().getString());
			}
			break;
		default:
			field.set(obj, cell.getStringCellValue());
			break;
		}
	}

	/**
	 * 导出Excel字节数据
	 *
	 * @param dataList
	 * @return
	 */
	public static byte[] exportToBytes(List<?>... dataArrs) {
		return exportToBytes(null, dataArrs);
	}

	/**
	 * 导出Excel字节数据
	 * 
	 * @param suffix
	 *            文件后缀xls或xlsx
	 * @param dataArrs
	 * @return
	 */
	public static byte[] exportToBytes(String suffix, List<?>... dataArrs) {
		if (suffix == null || "".equals(suffix.trim())) {
			suffix = XLSX;
		}
		// workbook
		Workbook workbook = createWorkbook(null, suffix, dataArrs);

		ByteArrayOutputStream byteArrayOutputStream = null;
		byte[] result = null;
		try {
			// workbook 2 ByteArrayOutputStream
			byteArrayOutputStream = new ByteArrayOutputStream();
			workbook.write(byteArrayOutputStream);

			// flush
			byteArrayOutputStream.flush();

			result = byteArrayOutputStream.toByteArray();
			return result;
		} catch (Exception e) {
			logger.error(e.getMessage(), e);
			throw new RuntimeException(e);
		} finally {
			try {
				if (byteArrayOutputStream != null) {
					byteArrayOutputStream.close();
				}
				if (workbook != null) {
					workbook.close();
				}
			} catch (Exception e) {
				logger.error(e.getMessage(), e);
				throw new RuntimeException(e);
			}
		}
	}

	/**
	 * 默认为2007版excel，兼容2003版excel
	 * 
	 * @param filePath
	 * @param datas
	 */
	public static void createExcel(String filePath, List<?>... datas) {
		createExcel(filePath, null, datas);
	}

	/**
	 * 默认为2007版excel，兼容2003版excel
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param datas
	 */
	public static void createExcel(String filePath, String sheetName, List<?>... datas) {
		String suffix = XLSX;
		if (filePath != null && !"".equals(filePath.trim())) {
			suffix = filePath.substring(filePath.lastIndexOf(".") + 1);
		}
		// workbook
		Workbook workbook = createWorkbook(sheetName, suffix, datas);

		FileOutputStream fileOutputStream = null;
		try {
			// workbook 2 FileOutputStream
			fileOutputStream = new FileOutputStream(filePath);
			workbook.write(fileOutputStream);

			// flush
			fileOutputStream.flush();
		} catch (Exception e) {
			logger.error(e.getMessage(), e);
			throw new RuntimeException(e);
		} finally {
			try {
				if (fileOutputStream != null) {
					fileOutputStream.close();
				}
				if (workbook != null) {
					workbook.close();
				}
			} catch (Exception e) {
				logger.error(e.getMessage(), e);
				throw new RuntimeException(e);
			}
		}
	}

	/**
	 * 
	 * @param suffix
	 *            文件后缀
	 * @param datas
	 * @return
	 */
	private static Workbook createWorkbook(String sheetName, String suffix, List<?>... dataArrs) {
		if (dataArrs == null || dataArrs.length == 0) {
			throw new RuntimeException("-----> fast-excel error, datas can not be empty.");
		}

		// book
		Workbook workbook = null;
		// HSSFWorkbook=2003/xls、XSSFWorkbook=2007/xlsx
		if (suffix != null && suffix.equals(XLS)) {
			workbook = new HSSFWorkbook(); // 2003
		} else {
			workbook = new XSSFWorkbook(); // 2007
		}

		// sheet
		for (List<?> datas : dataArrs) {
			createSheet(sheetName, workbook, datas);
		}

		return workbook;
	}

	private static void createSheet(String sheetName, Workbook workbook, List<?> datas) {
		if (datas == null || datas.size() == 0) {
			throw new RuntimeException("------> Fast-excel error, data can not be empty.");
		}

		// sheet
		Class<?> clazz = datas.get(0).getClass();

		HSSFColor.HSSFColorPredefined headColor = null;
		if (clazz.isAnnotationPresent(ExcelSheet.class)) {
			ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
			if (sheetName == null || "".equals(sheetName.trim())) {
				if (excelSheet.name() != null && !"".equals(excelSheet.name().trim())) {
					sheetName = excelSheet.name().trim();
				} else if (excelSheet.value() != null && !"".equals(excelSheet.value().trim())) {
					sheetName = excelSheet.value().trim();
				} else {
					sheetName = "sheet 1";
				}
			}
			headColor = excelSheet.headColor();
		}

		Sheet existSheet = workbook.getSheet(sheetName);
		if (existSheet != null) {
			for (int i = 2; i <= 1000; i++) {
				// 避免sheetName重复
				String newSheetName = sheetName.concat(String.valueOf(i));

				existSheet = workbook.getSheet(newSheetName);
				if (existSheet == null) {
					sheetName = newSheetName;
					break;
				} else {
					continue;
				}
			}
		}

		Sheet sheet = workbook.createSheet(sheetName);

		// sheet field
		Map<String, Field> fieldMap = new HashMap<String, Field>();
		Map<Integer, String> snMap = new TreeMap<Integer, String>();
		Map<Integer, String[]> dataValidityMap = new TreeMap<Integer, String[]>();
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			if (Modifier.isStatic(field.getModifiers())) {
				continue;
			}
			if (field.isAnnotationPresent(ExcelCell.class)) {
				ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
				String cellName = "";
				if (excelCell.name() != null && !"".equals(excelCell.name().trim())) {
					cellName = excelCell.name();
				} else if (excelCell.value() != null && !"".equals(excelCell.value().trim())) {
					cellName = excelCell.value();
				} else {
					cellName = field.getName();
				}
				fieldMap.put(cellName, field);
				snMap.put(excelCell.sn(), cellName);
				if (excelCell.data_validity().length != 0) {
					dataValidityMap.put(excelCell.sn(), excelCell.data_validity());
				}
			}
		}

		if (fieldMap.isEmpty()) {
			throw new RuntimeException("---------> Fast-excel error, data field has not been annotation.");
		}
		
        // sheet header row
		CellStyle headStyle = null;
		if (headColor != null) {
			headStyle = workbook.createCellStyle();
            /*Font headFont = book.createFont();
            headFont.setColor(headColor);
            headStyle.setFont(headFont);*/

			headStyle.setFillForegroundColor(headColor.getIndex());
			headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headStyle.setFillBackgroundColor(headColor.getIndex());
		}

		Collection<String> values = snMap.values();
		String[] s = new String[values.size()];
		values.toArray(s);
		// 生成标题行
		Row headRow = sheet.createRow(0);
		for (int i = 0; i < s.length; i++) {
			Cell cellX = headRow.createCell(i, CellType.STRING);
			if (headStyle != null) {
				cellX.setCellStyle(headStyle);
			}
			cellX.setCellValue(s[i]);
		}

		// 生成数据行（sheet data rows）
		for (int dataIndex = 0; dataIndex < datas.size(); dataIndex++) {
			int rowIndex = dataIndex + 1;
			Object rowData = datas.get(dataIndex);

			Row rowX = sheet.createRow(rowIndex);

			for (int j = 0; j < s.length; j++) {
				Cell cell = rowX.createCell(j, CellType.STRING);
				for (Map.Entry<String, Field> data : fieldMap.entrySet()) {
					try {
						if (data.getKey().equals(s[j])) {
							Field field = data.getValue();
							field.setAccessible(true);
							Object fieldValue = field.get(rowData);
							if (fieldValue == null) {
								cell.setCellValue("");
							} else {
								String fieldValueString = formatValue(field, fieldValue);
								cell.setCellValue(fieldValueString);
							}
							break;
						}
					} catch (SecurityException | IllegalArgumentException | IllegalAccessException e) {
						logger.error(e.getMessage(), e);
						throw new RuntimeException(e);
					}
				}
			}
		}

		// 添加数据有效性
		if (dataValidityMap.size() != 0) {
			DataValidationHelper helper = sheet.getDataValidationHelper();
			for (int col : dataValidityMap.keySet()) {
				// 设置一个需要提供下拉的区域
				String[] listValidity = dataValidityMap.get(col);
				// 确定下拉列表框的位置
				CellRangeAddressList regions = new CellRangeAddressList(1, 9000, col, col);
				// 生成下拉列表框的内容
				DataValidationConstraint constraint = helper.createExplicitListConstraint(listValidity);
				// 绑定下拉框的作用区域
				DataValidation dataValidation = helper.createValidation(constraint, regions);

				// 处理Excel兼容性问题
				if (dataValidation instanceof XSSFDataValidation) {
					dataValidation.setSuppressDropDownArrow(true);
					dataValidation.setShowErrorBox(true);
					dataValidation.setEmptyCellAllowed(false);
				} else {
					dataValidation.setSuppressDropDownArrow(false);
					dataValidation.setEmptyCellAllowed(false);
				}
				// 对哪一页起作用
				sheet.addValidationData(dataValidation);
			}
		}

	}
	
	/**
	 * 参数格式化为String
	 *
	 * @param field
	 * @param value
	 * @return
	 */
	private static String formatValue(Field field, Object value) {
		Class<?> fieldType = field.getType();

		ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
		if (value == null) {
			return null;
		}

		if (Boolean.class.equals(fieldType) || Boolean.TYPE.equals(fieldType)) {
			return String.valueOf(value);
		} else if (String.class.equals(fieldType)) {
			return String.valueOf(value);
		} else if (Short.class.equals(fieldType) || Short.TYPE.equals(fieldType)) {
			return String.valueOf(value);
		} else if (Integer.class.equals(fieldType) || Integer.TYPE.equals(fieldType)) {
			return String.valueOf(value);
		} else if (Long.class.equals(fieldType) || Long.TYPE.equals(fieldType)) {
			return String.valueOf(value);
		} else if (Float.class.equals(fieldType) || Float.TYPE.equals(fieldType)) {
			return String.valueOf(value);
		} else if (Double.class.equals(fieldType) || Double.TYPE.equals(fieldType)) {
			return String.valueOf(value);
		} else if (Date.class.equals(fieldType)) {
			String datePattern = "yyyy-MM-dd HH:mm:ss";
			if (excelCell != null && excelCell.dateformat() != null && !"".equals(excelCell.dateformat().trim())) {
				datePattern = excelCell.dateformat();
			}
			SimpleDateFormat dateFormat = new SimpleDateFormat(datePattern);
			return dateFormat.format(value);
		} else {
			throw new RuntimeException(
					"request illeagal type, type must be Integer not int Long not long etc, type=" + fieldType);
		}
	}

	/**
	 * 获取指定单元格的值
	 * 
	 * @param workbook
	 * @param sheetName
	 * @param rowNumber
	 *            行数，从1开始
	 * @param cellNumber
	 *            列数，从1开始
	 * @return 该单元格的值
	 */
	public String getCellValue(Workbook workbook, String sheetName, int rowNumber, int cellNumber) {
		String result = null;
		checkRowAndCell(rowNumber, cellNumber);
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.getRow(--rowNumber);
		Cell cell = row.getCell(--cellNumber);
		switch (cell.getCellTypeEnum()) {
		case BLANK:
			result = cell.getStringCellValue();
			break;
		case BOOLEAN:
			result = String.valueOf(cell.getBooleanCellValue());
			break;
		case ERROR:
			result = String.valueOf(cell.getErrorCellValue());
			break;
		case FORMULA:
			result = cell.getCellFormula();
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				result = dateFormat.format(cell.getDateCellValue());
			} else {
				result = String.valueOf(cell.getNumericCellValue());
			}
			break;
		case STRING:
			result = cell.getRichStringCellValue().getString();
			break;
		default:
			result = cell.getStringCellValue();
			break;
		}
		return result;
	}

	private void checkRowAndCell(int rowNumber, int cellNumber) {
		if (rowNumber < 1) {
			throw new RuntimeException("rowNumber less than 1");
		}
		if (cellNumber < 1) {
			throw new RuntimeException("cellNumber less than 1");
		}
	}

}
