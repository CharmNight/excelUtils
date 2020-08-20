package com.night.excel.beans;

import java.io.File;
import java.lang.reflect.Field;
import java.lang.reflect.Method;

/**
 * excel 映射关系
 *
 * @Author: CharmNight
 * @Date: 2020/8/21 1:19
 */
public class ExcelExportEntity {
    private String name;
    private Field field;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }
}
