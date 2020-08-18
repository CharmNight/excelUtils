package com.night.excel.utils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * @Author: CharmNight
 * @Date: 2020/8/19 0:27
 */
public class ReflectionUtils {

    /**
     * 通过反射获取get方法的返回值
     * @param field
     * @param clazz
     * @param insert
     * @return
     */
    public static Object getMethodRes(Field field, Class clazz, Object insert){
        field.setAccessible(true);
        String name = field.getName();
        name = name.substring(0,1).toUpperCase() + name.substring(1, name.length());
        String setMethodName = "get" + name;

        Method method = null;

        try {
            Class<Object>[] argsClass = new Class[0];

            method = clazz.getMethod(setMethodName, argsClass);

            Object res = method.invoke(insert);

            return res;

        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }
        return new Object();
    }
}
