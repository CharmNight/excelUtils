package com.night.excel.beans;


import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.ArrayList;
import java.util.List;

/**
 * @Author: CharmNight
 * @Date: 2020/8/19 0:14
 */

@SpringBootTest
class UserInfoExcelBeanTest {
    @Autowired
    private UserInfoExcelBean bean;

    @Test
    public void test(){

        System.out.println(bean.getExcelTitles());
        System.out.println(bean.getExcelName());

        List<UserInfoExcelBean> list = new ArrayList<>();

        for (int i = 0; i < 65540; i++) {
            UserInfoExcelBean userInfoExcelBean = new UserInfoExcelBean();
            userInfoExcelBean.setName("CharmNight");
            userInfoExcelBean.setCreateUser("CharmNight" + i);
            userInfoExcelBean.setType("up");
            userInfoExcelBean.setUpdatedAt("2012-12-12 12:12:12");
            list.add(userInfoExcelBean);
        }
        bean.exportToExcel("D:\\JavaCode\\pubGit\\excelUtils\\excel\\src\\main\\java\\com\\night\\excelTestFiles\\user2.xls", list);

    }

}