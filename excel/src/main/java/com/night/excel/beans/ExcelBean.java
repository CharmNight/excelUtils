package com.night.excel.beans;

import com.night.excel.fields.ExcelName;
import com.night.excel.fields.ExcelTitle;
import com.night.excel.utils.ExcelUtils;
import com.night.excel.utils.ReturnClient;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ExcelBean<T> {
	private String name = "";
	private List<ExcelExportEntity> title = new ArrayList<>();

	@Autowired
	private ExcelUtils utils;

	public ExcelBean() {
		setExcelTitles();
		setExcelName();
	}

	private void setExcelTitles(){
		Field[] fields = this.getClass().getDeclaredFields();
		for (Field field : fields) {
			ExcelTitle excelTitle = field.getAnnotation(ExcelTitle.class);
			ExcelExportEntity excelExportEntity = new ExcelExportEntity();
			excelExportEntity.setName(excelTitle.value());
			excelExportEntity.setField(field);
			this.title.add(excelExportEntity);
		}
	}

	private void setExcelName(){
		ExcelName excelName = this.getClass().getAnnotation(ExcelName.class);
		String name = excelName.value();
		this.name = name;
	}

	public List<ExcelExportEntity> getExcelTitles() {
		return this.title;
	}

	public String getExcelName() {
		return this.name;
	}

	/**
	 * 导出本地
	 *
	 * @param path
	 * @param list
	 * @return
	 */
	public HSSFWorkbook exportToExcel(String path, List<T> list){
		try {
			ReturnClient.returnClient(path, utils.pubGetHSSFWorkbook(name, title, list));
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * 导出本地, 支持 同时写入多个sheet页,
	 * 如果 sheet 达到最大值 默认65535 可以自动扩展页 自动新增sheet名称规则: sheet_1\2\3
	 * @param path
	 * @param list
	 * @return
	 */
	public HSSFWorkbook appendExportToExcel(String path, Map<String, List<T>> list){
		try {
			HSSFWorkbook wb = utils.createHSSFWorkbook(null);
			for (Map.Entry<String, List<T>> entry : list.entrySet()) {
				String sheetName = entry.getKey();
				List<T> value = entry.getValue();
				utils.appendData(wb, sheetName, title, value);
			}
			ReturnClient.returnClient(path, wb);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}


}
