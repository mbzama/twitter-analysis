package zama.learning.twitter.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import zama.learning.twitter.excel.annotation.ExcelFilter;
import zama.learning.twitter.excel.annotation.ExcelGrid;

public class ExcelGenerator<T> {
	private static final Logger LOGGER = Logger.getLogger(ExcelGenerator.class);
	private static final int NUMBER_OF_CELLS_PER_LINE = 8;
	private int contentStartRow = 0;

	public HSSFWorkbook generateExcelWithForm(List<T> data, Object form, List<List<String>> rows, List<Integer> excludedcolumns) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFFont font = getnormalFont(workbook);
		int rowNum = 0;

		HSSFSheet sheet = workbook.createSheet();

		if (form != null) {
			rowNum = createSelectionCriteria(form, sheet, workbook, font);
		}

		contentStartRow = rowNum;

		int lastRow = generateWorkbook(workbook, data, rowNum, excludedcolumns);

		if (rows != null && rows.size() > 0) {
			lastRow++;

			HSSFCellStyle white = createEvenCellStyle(workbook);
			HSSFCellStyle odd = createOddCellStyle(workbook);

			createExtraRows(workbook, white, odd, lastRow, rows);
		}

		return workbook;
	}

	public HSSFWorkbook generateExcel(List<T> data) {
		return this.generateExcelWithForm(data, null, null, null);
	}

	public HSSFWorkbook generateExcel(List<T> data, List<List<String>> rows, List<Integer> excludedcolumns) {
		return this.generateExcelWithForm(data, null, rows, excludedcolumns);
	}

	public HSSFWorkbook generateExcel(HSSFWorkbook workbook, List<T> data) {
		HSSFSheet sheet = workbook.getSheetAt(0);
		contentStartRow = sheet.getLastRowNum() + 1;
		generateWorkbook(workbook, data, contentStartRow, null);

		return workbook;
	}

	public HSSFWorkbook generateExcel(HSSFWorkbook workbook, List<T> data, List<Integer> excludedcolumns) {
		HSSFSheet sheet = workbook.getSheetAt(0);
		contentStartRow = sheet.getLastRowNum() + 1;
		generateWorkbook(workbook, data, contentStartRow, excludedcolumns);

		return workbook;
	}

	private void createExtraRows(HSSFWorkbook workbook,  HSSFCellStyle white, HSSFCellStyle odd, int rowCnt, List<List<String>> rows) {
		int rowNum = rowCnt;
		HSSFSheet sheet = workbook.getSheetAt(0);

		// Add on the extra rows sent in
		if (rows != null && rows.size() > 0) {
			rowNum++;
			boolean firstRow = true;

			HSSFCellStyle cellStyle = null;

			for (List<String> rowList : rows) {
				rowNum++;
				HSSFRow row = sheet.createRow(rowNum);
				int cellCnt = 0;

				for (String s : rowList) {
					HSSFCell cell = row.createCell(cellCnt, HSSFCell.CELL_TYPE_STRING);
					cellCnt++;

					// stripping for even and odd rows.
					if (firstRow) {
						cellStyle = createHeaderCellStyle(workbook);
						cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					} else {

						cellStyle = white;
					}

					cell.setCellStyle(cellStyle);
					cell.setCellValue(s);
				}

				firstRow = false;
			}
		}
	}

	private int generateWorkbook(HSSFWorkbook workbook, List<T> data, int rowNum, List<Integer> excludedcolumns) {
		HSSFFont font = getnormalFont(workbook);
		HSSFSheet sheet = workbook.getSheetAt(0);
		List<Method> excelInfo = new ArrayList<Method>();
		int rowCnt = rowNum;

		HSSFRow row = sheet.createRow(rowCnt);
		rowCnt++;
		T headerRow = data.get(0);
		LOGGER.debug("header: " + headerRow);
		LOGGER.debug("class: " + headerRow.getClass());
		Method methods[] = headerRow.getClass().getDeclaredMethods();
		List<Method> methodsList = new ArrayList<Method>();
		for (Method method : methods) {
			ExcelGrid excelHeader = method.getAnnotation(ExcelGrid.class);
			if(excelHeader!=null){
				methodsList.add(method);
			}
		}

		Collections.sort(methodsList, new Comparator<Method>() {

			@Override
			public int compare(Method o1, Method o2) {
				ExcelGrid excelHeader1 = o1.getAnnotation(ExcelGrid.class);
				ExcelGrid excelHeader2 = o2.getAnnotation(ExcelGrid.class);
				int val = 0;
				if(excelHeader1.order()==excelHeader2.order()){
					val = 0;
				}else if(excelHeader1.order()>excelHeader2.order()){
					val = 2;
				}else if(excelHeader1.order()<excelHeader2.order()){
					val = -1;
				}
				return val;
			}
		});

		for (Method method : methodsList) {
			ExcelGrid excelHeader = method.getAnnotation(ExcelGrid.class);
			boolean exclude = false;

			if (excelHeader != null) {
				if (excludedcolumns != null && excludedcolumns.size() > 0) {
					for (int i : excludedcolumns) {
						if (i == excelHeader.order()) {
							exclude = true;
						}
					}
				}
				LOGGER.debug("order: " + excelHeader.order());
				LOGGER.debug("align: " + excelHeader.align());
				LOGGER.debug("header: " + excelHeader.header());
				if (!exclude) {
					excelInfo.add(method);
				}
			}
		}
		createHeader(row, workbook, excelInfo);
		HSSFCellStyle even = createEvenCellStyle(workbook);
		HSSFCellStyle odd = createOddCellStyle(workbook);

		for (T item : data) {
			row = sheet.createRow(rowCnt);
			rowCnt++;

			generateRow(row, workbook, item, font, excelInfo, odd, even);
		}

		this.autoSizeColumns(sheet);
		return rowCnt;

	}

	private int createSelectionCriteria(Object form, HSSFSheet sheet, HSSFWorkbook workbook, HSSFFont normalFont) {
		int rowNum = 0;

		int cellNum = 0;
		HSSFRow row = sheet.createRow(rowNum);

		HSSFFont boldFont = getBoldFont(workbook);
		HSSFCellStyle hdrStyle = createHeaderCellStyle(workbook);
		hdrStyle.setFont(boldFont);

		HSSFCell headerRowCell = row.createCell(0);
		hdrStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		headerRowCell.setCellStyle(hdrStyle);
		headerRowCell.setCellValue("Selection Criteria");
		CellRangeAddress row0Range = new CellRangeAddress(0, 0, 1, 11);
		sheet.addMergedRegion(row0Range);

		// The date for the report date
		String reportDate = new Date().toString();
		HSSFCell dateCell = row.createCell(1);
		dateCell.setCellStyle(hdrStyle);
		dateCell.setCellValue(reportDate);

		rowNum++;

		for (Method method : form.getClass().getDeclaredMethods()) {
			ExcelFilter excelFilter = method.getAnnotation(ExcelFilter.class);

			if (excelFilter != null) {

				//**************** set-up the row to use
				if (excelFilter.row() > 0) {
					if (excelFilter.row() > rowNum) {
						rowNum = excelFilter.row();
					}
					row = sheet.getRow(excelFilter.row());
					if (row == null) {
						row = sheet.createRow(excelFilter.row());
					}
				} else {
					row = sheet.getRow(rowNum);
					if (row == null) {
						row = sheet.createRow(rowNum);
					}
				}

				//**************** set-up the cell to use
				HSSFCell cell = null;

				if (excelFilter.order() > 0) {
					cellNum = (excelFilter.order() -1) * 2;
				}

				cell = row.getCell(cellNum, Row.CREATE_NULL_AS_BLANK);

				Object item = null;
				try {
					item = (Object)method.invoke(form);
				} catch (IllegalArgumentException e) {
					LOGGER.warn("illegal argument: " + e);
				} catch (IllegalAccessException e) {
					LOGGER.warn("illegal Access: " + e);
				} catch (InvocationTargetException e) {
					LOGGER.warn("invocation target exception: " + e);
				}
				String value = "";

				if (item != null) {
					if (item.getClass().isArray()) {
						value = arrayToString((Object[]) item);
					} else {
						value = item.toString();
					}
				}

				if (cellNum > NUMBER_OF_CELLS_PER_LINE && excelFilter.row() == 0) {
					row = sheet.createRow(rowNum);
					cellNum = 0;
					rowNum++;
					cell = row.createCell(cellNum);
				}

				//*********** set-up cell style
				HSSFCellStyle style = createEvenCellStyle(workbook);
				style.setFont(boldFont);
				style.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
				cell.setCellValue(excelFilter.title());
				cell.setCellStyle(style);

				cellNum++;
				cell = row.createCell(cellNum);
				style = createEvenCellStyle(workbook);
				style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
				cell.setCellValue(value);
				cellNum++;
				cell.setCellStyle(style);
				LOGGER.debug("title:value =" + excelFilter.title() + ":" + value);
			}
		}

		rowNum++;

		return rowNum;
	}

	private String arrayToString (Object[] arr) {
		String rtn = "";
		int count = 0;

		for (Object obj : arr) {
			rtn += (count == 0) ? obj.toString() : ", " + obj.toString();
			count++;
		}

		return rtn;
	}

	private void createHeader(HSSFRow row, HSSFWorkbook workbook, List<Method> data) {
		int currentCell = 0;
		int i = 0;
		for (Method method : data) {
			ExcelGrid excelGrid = method.getAnnotation(ExcelGrid.class);

			HSSFCell cell = null;

			if (excelGrid.order() > 0) {
				cell = row.getCell((i), Row.CREATE_NULL_AS_BLANK);
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			} else {
				cell = row.createCell(currentCell, HSSFCell.CELL_TYPE_STRING);
				currentCell++;
			}
			i++;
			HSSFCellStyle cellStyle = createHeaderCellStyle(workbook);

			cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

			cell.setCellStyle(cellStyle);

			cell.setCellValue(excelGrid.header());
			LOGGER.debug("header: " + excelGrid.header() + " order:" + excelGrid.order());
		}
	}

	private void generateRow(HSSFRow row, HSSFWorkbook workbook, T data, HSSFFont font, 
			List<Method> methods, HSSFCellStyle odd, HSSFCellStyle even) {
		int currentCell = 0;
		int i = 0;
		// spin through all the methods that are annotated and invoke them to get the cell values.
		for (Method method : methods) {
			ExcelGrid excelGrid = method.getAnnotation(ExcelGrid.class);
			short align = excelGrid.align();
			Object item; 

			try {
				item = (Object)method.invoke(data);
				String value = "";
				if (item != null) {
					value = item.toString();
				}
				HSSFCell cell = null;

				if (excelGrid.order() > 0) {
					cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
					cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				} else {
					cell = row.createCell(currentCell, HSSFCell.CELL_TYPE_STRING);
					currentCell++;
				}
				i++;
				// stripping for even and odd rows.
				HSSFCellStyle cellStyle = null;
				if (row.getRowNum() % 2 == 1) {
					cellStyle = odd;
				} else {
					cellStyle = even;
				}

				cellStyle.setFont(font);
				cellStyle.setAlignment(align);
				cell.setCellStyle(cellStyle);

				cell.setCellValue(value);
			} catch (IllegalArgumentException e) {
				LOGGER.debug("illegal arg" + e);
			} catch (IllegalAccessException e) {
				LOGGER.debug("illegal access" + e);
			} catch (InvocationTargetException e) {
				LOGGER.debug("invoc target bad" + e);
			} catch (SecurityException e) {
				LOGGER.debug("Security Exception: " + e);
			}
		}
	}

	private HSSFCellStyle createHeaderCellStyle(HSSFWorkbook workbook) {
		HSSFCellStyle cellStyle = createEvenCellStyle(workbook);

		//HSSFPalette palette = workbook.getCustomPalette();
		//palette.setColorAtIndex(HSSFColor.LIGHT_TURQUOISE.index, (byte) 100, (byte) 175, (byte) 63);
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellStyle.setFillForegroundColor(HSSFColor.AQUA.index);

		HSSFFont font = workbook.createFont();
		font.setFontName("Calibri");
		font.setColor(HSSFFont.COLOR_NORMAL);
		short fontSize = 10;
		font.setFontHeightInPoints(fontSize);
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

		cellStyle.setFont(font);

		return cellStyle;
	}

	private HSSFFont getnormalFont(HSSFWorkbook workbook) {
		HSSFFont font = workbook.createFont();

		font.setFontName("Calibri");
		font.setColor(HSSFFont.COLOR_NORMAL);
		short fontSize = 10;
		font.setFontHeightInPoints(fontSize);

		return font;
	}

	private HSSFFont getBoldFont(HSSFWorkbook workbook) {
		HSSFFont font = workbook.createFont();

		font.setFontName("Calibri");
		font.setColor(HSSFFont.COLOR_NORMAL);
		short fontSize = 10;
		font.setFontHeightInPoints(fontSize);
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);

		return font;
	}

	private HSSFCellStyle createOddCellStyle(HSSFWorkbook workbook) {
		HSSFCellStyle cellStyle = createEvenCellStyle(workbook);

		//HSSFPalette palette = workbook.getCustomPalette();
		//palette.setColorAtIndex(HSSFColor.AQUA.index, (byte) 239, (byte) 232, (byte) 210);
		//cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		//cellStyle.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);

		return cellStyle;
	}

	private HSSFCellStyle createEvenCellStyle(HSSFWorkbook workbook) {
		HSSFCellStyle cellStyle = workbook.createCellStyle();

		return cellStyle;
	}

	private void autoSizeColumns(HSSFSheet sheet) {
		HSSFRow row = sheet.getRow(contentStartRow);
		if (row != null) {
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				sheet.autoSizeColumn(j);
			}
		}
	}

	public static void writeFile(HSSFWorkbook workbook, File file) {
		LOGGER.debug("writing out the file");
		FileOutputStream fileOutputStream = null;
		try {
			fileOutputStream = new FileOutputStream(file);
			workbook.write(fileOutputStream);
		} catch (FileNotFoundException e) {
			LOGGER.error("File not found: " + e);
		} catch (IOException e) {
			LOGGER.error("IO Exception: " + e);
		} finally {
			if (fileOutputStream != null) {
				try {
					fileOutputStream.close();
				} catch (IOException e) {
					LOGGER.error("IO Exception when trying to close the file: " + e);
				}
			}
		}
	}

}
