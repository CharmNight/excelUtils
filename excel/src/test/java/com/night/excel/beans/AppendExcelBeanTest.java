package com.night.excel.beans;


import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author: CharmNight
 * @Date: 2020/8/19 0:14
 */

@SpringBootTest
class AppendExcelBeanTest {
    @Autowired
    private UserInfoExcelBean bean;

//    @Test
//    public void test(){
//
//        System.out.println(bean.getExcelTitles());
//        System.out.println(bean.getExcelName());
//
//        Map<String, List<UserInfoExcelBean>> map = new HashMap<>();
//
//        for (int j = 0; j < 3; j++) {
//            String sheetName = "Sheet" + "2020" + j;
//            List<UserInfoExcelBean> list = new ArrayList<>();
//            for (int i = 0; i < 10; i++) {
//                UserInfoExcelBean userInfoExcelBean = new UserInfoExcelBean();
//                userInfoExcelBean.setName("CharmNight");
//                userInfoExcelBean.setCreateUser("CharmNight" + i);
//                userInfoExcelBean.setType("up");
//                userInfoExcelBean.setUpdatedAt("2012-12-12 12:12:12");
//                list.add(userInfoExcelBean);
//            }
//            map.put(sheetName, list);
//        }
//
//        bean.appendExportToExcel("D:\\JavaCode\\pubGit\\excelUtils\\excel\\src\\main\\java\\com\\night\\excelTestFiles\\user3.xls", map);
//
//    }
    @Test
    public void test2(){

        System.out.println(bean.getExcelTitles());
        System.out.println(bean.getExcelName());

        // Map <String: sheetName, List<T>: 列的值>
        Map<String, List<UserInfoExcelBean>> map = new HashMap<>();

        for (int j = 0; j < 3; j++) {
            String sheetName = "Sheet" + "2020" + j;
            List<UserInfoExcelBean> list = new ArrayList<>();
            for (int i = 0; i < 65540; i++) {
                UserInfoExcelBean userInfoExcelBean = new UserInfoExcelBean();
                userInfoExcelBean.setName("CharmNight");
                userInfoExcelBean.setCreateUser("CharmNight" + i);
                userInfoExcelBean.setType("up");
                userInfoExcelBean.setUpdatedAt("2012-12-12 12:12:12");
                list.add(userInfoExcelBean);
            }
            map.put(sheetName, list);
        }

        bean.appendExportToExcel("D:\\JavaCode\\pubGit\\excelUtils\\excel\\src\\main\\java\\com\\night\\excelTestFiles\\user3.xls", map);

    }

}