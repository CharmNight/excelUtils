package com.night.excel.utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Controller;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Objects;

@Component
public class ExcelUtils<T> {
	/**
	 * Excel 每页最大行数限制
	 */
	private static final int MAX_LENS = 65535;


	public HSSFWorkbook pubGetHSSFWorkbook(String sheetName, List<String> title, List<T> values) {
		HSSFWorkbook wb = getHSSFWorkbook(sheetName, title, null);
		HSSFSheet sheet = getCurrentSheet(wb);
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

		return wb;
	}

	/**
	 * 获取 HSSFWorkbook
	 * @param sheetName
	 * @param sheetPage
	 * @param title
	 * @param wb
	 * @return
	 */
	private HSSFWorkbook getHSSFWorkbook(String sheetName, Integer sheetPage, List<String> title, HSSFWorkbook wb) {
		// 第一步，创建一个HSSFWorkbook，对应一个Excel文件
		if (wb == null) {
			wb = new HSSFWorkbook();
		}
//		getSheet(sheetName, title, sheetPage, wb);
		return wb;
	}

	/**
	 * 获取 HSSFWorkbook
	 * @param sheetName
	 * @param title
	 * @param wb
	 * @return
	 */
	private HSSFWorkbook getHSSFWorkbook(String sheetName, List<String> title, HSSFWorkbook wb){
		// get wb
		if (Objects.isNull(wb)) {
			wb = getHSSFWorkbook(sheetName, 0, title, null);
		}
		return wb;
	}

	/**
	 * 获取页
	 * @param wb
	 * @param sheetPage
	 * @return
	 */
	private HSSFSheet getSheet(HSSFWorkbook wb, int sheetPage) {
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
	private HSSFSheet getSheet(String sheetName, List<String> title, Integer sheetPage, HSSFWorkbook wb) {
		if (sheetPage !=0) {
			sheetName = sheetName + sheetPage;
		}
		HSSFSheet sheet = wb.getSheet(sheetName);
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
	private void setTitle(List<String> title, HSSFSheet sheet, HSSFWorkbook wb) {
		//声明列对象
		HSSFCell cell = null;
		HSSFRow row = sheet.createRow(0);
		// 第四步，创建单元格，并设置值表头 设置表头居中
		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
		//创建标题
		for (int i = 0; i < title.size(); i++) {
			cell = row.createCell(i);
			cell.setCellValue(title.get(i));
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
	private HSSFWorkbook writeContent(HSSFSheet sheet, List<T> values, HSSFWorkbook wb){

		int rows = sheet.getPhysicalNumberOfRows();
		//创建内容
		for (int i = 0; i < values.size(); i++) {
			HSSFRow row = sheet.createRow(rows + i);
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
	private HSSFSheet getCurrentSheet(HSSFWorkbook wb){
		// 获取当前页
		int sheetPage = getCurrentSheetNum(wb);
		HSSFSheet sheet = getSheet(wb, sheetPage);
		return sheet;

	}

	/**
	 * 获取当前页数
	 * @param wb
	 * @return
	 */
	private int getCurrentSheetNum(HSSFWorkbook wb) {
		int numberOfSheets = wb.getNumberOfSheets();
		return numberOfSheets != 0 ? numberOfSheets-1 : 0;
	}

	/**
	 * 判断是否需要扩展页
	 * @param sheet
	 * @param sheetPage
	 * @return true 需要
	 */
	private Boolean checkAddNewSheet(HSSFSheet sheet, int sheetPage){
		// 判断是否需要扩页
		if (sheetPage == 0) {
			if (Objects.isNull(sheet)) {
				return true;
			}
		}
		return false;
	}

}
