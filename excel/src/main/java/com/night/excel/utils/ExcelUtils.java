package com.night.excel.utils;

import com.night.excel.beans.ExcelExportEntity;
import com.night.excel.enmus.ExcelType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Controller;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

@Component
public class ExcelUtils<T> {
	/**
	 * Excel 每页最大行数限制
	 */
	private static final int MAX_LENS = 65535;

	private static final String SHEET_NAME_SPLIT = "_";

	public HSSFWorkbook pubGetHSSFWorkbook(String sheetName, List<ExcelExportEntity> title, List<T> values) {
		Workbook wb = getWorkbook(null, ExcelType.HSSF);
		Sheet sheet = getCurrentSheet(wb);
		int currentSheetNum = getCurrentSheetNum(wb);
		Boolean needNewSheet = checkAddNewSheet(sheet, 0);
		if (needNewSheet) {
			sheet = getSheet(sheetName, title, currentSheetNum, wb);
		}

		int rowsNum = sheet.getPhysicalNumberOfRows();

		if (rowsNum + values.size() >= MAX_LENS) {
			int pageSize = (rowsNum + values.size()) / MAX_LENS + 1;
			int startNum = 0;
			int startPage = 0;
			for (int i = startNum; i < pageSize; i++) {
				if (startPage == 0) {
					startPage = i * MAX_LENS;
				}

				int len = 0;
				if (i ==0 && (rowsNum != 1 && rowsNum != 0)) {
					// 非新开页数 查询当前页可写入行数
					len = MAX_LENS - rowsNum + 1;
					startPage = startPage != 0 ? len : 0;
				}else {
					len = i == pageSize -1 ? values.size() % MAX_LENS : MAX_LENS-rowsNum+1;

				}

				List<T> newValues = values.subList(startPage, startPage+len);
				sheet = getSheet(sheetName, title, currentSheetNum, wb);
				int rows = sheet.getPhysicalNumberOfRows() - 1;
				if((rows + newValues.size()) > MAX_LENS){
					newValues = values.subList(startPage, MAX_LENS - rows);
					startNum = MAX_LENS - rows;
				}
				writeContent(sheet, newValues, wb);
				currentSheetNum += 1;
				startPage = 0;
			}
		}else {
			writeContent(sheet, values, wb);
		}

		return (HSSFWorkbook) wb;
	}

	public Sheet createdSheet(Workbook wb, String sheetName){
		try {
			return wb.createSheet(sheetName);
		}catch (Exception e){
			// 这个名字重复
			String baseSheetName = sheetName.split(SHEET_NAME_SPLIT)[0];
			for (int i = 1; ; i++) {
				try {
					return wb.createSheet(baseSheetName + SHEET_NAME_SPLIT + i);
				}catch (Exception e2){
					continue;
				}
			}
		}
	}

	/**
	 *
	 * 另一种追加形式, 使用时 可以支持 写入多个sheet
	 *
	 * @param wb
	 * @param sheetName
	 * @param title
	 * @param values
	 * @return
	 */
	public Workbook appendData(Workbook wb, String sheetName, List<ExcelExportEntity> title, List<T> values) {
		HSSFWorkbook hssfWorkbook = createHSSFWorkbook(wb);

		Sheet sheet = createdSheet(wb, sheetName);
		// 表格样式暂无  QAQ

		// 当前仅支持 title 为1
		setTitle(title, sheet, hssfWorkbook);
		// 暂时 手动 title + 1
		int rowSize = sheet.getPhysicalNumberOfRows();

		Iterator<T> iterator = values.iterator();
		while (iterator.hasNext()){
			T next = iterator.next();

			if (rowSize > MAX_LENS) {
				sheet = createdSheet(wb, sheet.getSheetName());
				setTitle(title, sheet, hssfWorkbook);
				rowSize = sheet.getPhysicalNumberOfRows();
			}
			rowSize += createCells(rowSize, next, title, sheet, hssfWorkbook);
		}

		return wb;
	}

	public HSSFWorkbook createHSSFWorkbook(Workbook wb) {
		return (HSSFWorkbook) getWorkbook(wb, ExcelType.HSSF);

	}

	private int createCells(int rowSize, T next, List<ExcelExportEntity> fields, Sheet sheet, Workbook wb) {
		System.out.println(rowSize);
		Row row = sheet.getRow(rowSize) == null ? sheet.createRow(rowSize) : sheet.getRow(rowSize);
		for (int k = 0, paramSize = fields.size(); k < paramSize; k++){
			ExcelExportEntity excelExportEntity = fields.get(k);
			Field field = excelExportEntity.getField();
			field.setAccessible(true);
			try {
				 Object res = field.get(next);
				row.createCell(k).setCellValue(res.toString());
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			}
		}
		return 1;
	}


	/**
	 * 获取 Workbook
	 * @param wb
	 * @return
	 */
	private Workbook getWorkbook(Workbook wb, ExcelType excelType){
		// get wb
		if (Objects.isNull(wb)) {
			if (excelType.equals(ExcelType.HSSF)) {
				wb = new HSSFWorkbook();
			}
		}
		return wb;
	}

	/**
	 * 获取页
	 * @param wb
	 * @param sheetPage
	 * @return
	 */
	private Sheet getSheet(Workbook wb, int sheetPage) {
		try {
			return wb.getSheetAt(sheetPage);
		}catch (Exception e){
			return null;
		}
	}

	/**
	 * 获取页
	 * @param sheetName
	 * @param title
	 * @param sheetPage
	 * @param wb
	 * @return
	 */
	private Sheet getSheet(String sheetName, List<ExcelExportEntity> title, Integer sheetPage, Workbook wb) {
		if (sheetPage !=0) {
			sheetName = sheetName + sheetPage;
		}
		Sheet sheet = wb.getSheet(sheetName);
		if (Objects.isNull(sheet)) {
			sheet = wb.createSheet(sheetName);
			setTitle(title, sheet, wb);
		}
		return sheet;
	}

	/**
	 * 设置title
	 * @param title
	 * @param sheet
	 * @param wb
	 */
	private void setTitle(List<ExcelExportEntity> title, Sheet sheet, Workbook wb) {
		//声明列对象
		Cell cell = null;
		if (sheet.getPhysicalNumberOfRows() != 0) {
			return;
		}
		Row row = sheet.createRow(0);
		// 创建单元格，并设置值表头 设置表头居中
		CellStyle style = wb.createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
		//创建标题
		for (int i = 0; i < title.size(); i++) {
			cell = row.createCell(i);
			cell.setCellValue(title.get(i).getName());
			cell.setCellStyle(style);
		}
	}

	/**
	 * 写入内容
	 * @param sheet
	 * @param values
	 * @param wb
	 * @return
	 */
	private Workbook writeContent(Sheet sheet, List<T> values, Workbook wb){

		int rows = sheet.getPhysicalNumberOfRows();
		//创建内容
		for (int i = 0; i < values.size(); i++) {
			Row row = sheet.createRow(rows + i);
			T t = values.get(i);
			Class<?> clazz = t.getClass();

			Field[] fields = t.getClass().getDeclaredFields();
			for (int j = 0; j < fields.length; j++) {
				Field field = fields[j];
				Object res = ReflectionUtils.getMethodRes(field, clazz, t);
				//将内容按顺序赋给对应的列对象
				row.createCell(j).setCellValue(res.toString());
			}

		}
		return wb;
	}

	/**
	 * 获取最新页
	 * @param wb
	 * @return
	 */
	private Sheet getCurrentSheet(Workbook wb){
		// 获取当前页
		int sheetPage = getCurrentSheetNum(wb);
		Sheet sheet = getSheet(wb, sheetPage);
		return sheet;

	}

	/**
	 * 获取当前页数
	 * @param wb
	 * @return
	 */
	private int getCurrentSheetNum(Workbook wb) {
		int numberOfSheets = wb.getNumberOfSheets();
		return numberOfSheets != 0 ? numberOfSheets-1 : 0;
	}

	/**
	 * 判断是否需要扩展页
	 * @param sheet
	 * @param sheetPage
	 * @return true 需要
	 */
	private Boolean checkAddNewSheet(Sheet sheet, int sheetPage){
		// 判断是否需要扩页
		if (sheetPage == 0) {
			if (Objects.isNull(sheet)) {
				return true;
			}
		}
		return false;
	}

}
