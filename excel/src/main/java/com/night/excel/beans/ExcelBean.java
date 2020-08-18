package com.night.excel.beans;

import com.night.excel.fields.ExcelName;
import com.night.excel.fields.ExcelTitle;
import com.night.excel.utils.ExcelUtils;
import com.night.excel.utils.ReturnClient;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public class ExcelBean<T> {
	private String name = "";
	private List<String> title = new ArrayList<>();

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
			this.title.add(excelTitle.value());
		}
	}

	private void setExcelName(){
		ExcelName excelName = this.getClass().getAnnotation(ExcelName.class);
		String name = excelName.value();
		this.name = name;
	}

	public List<String> getExcelTitles() {
		return this.title;
	}

	public String getExcelName() {
		return this.name;
	}

	public HSSFWorkbook exportToExcel(String path, List<T> list){
		try {
			ReturnClient.returnClient(path, utils.pubGetHSSFWorkbook(name, title, list));
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}


}
