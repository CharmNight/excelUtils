package com.night.excel.beans;

import com.night.excel.beans.ExcelBean;
import com.night.excel.fields.ExcelName;
import com.night.excel.fields.ExcelTitle;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Controller;


/**
 * @Author: CharmNight
 * @Date: 2020/8/19 0:03
 */
@ExcelName(value="用户信息表")
@Component
public class UserInfoExcelBean extends ExcelBean<UserInfoExcelBean> {
    @ExcelTitle("用户名")
    private String name;
    @ExcelTitle("状态")
    private String type;
    @ExcelTitle("创建时间")
    private String updatedAt;
    @ExcelTitle("创建人")
    private String createUser;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getUpdatedAt() {
        return updatedAt;
    }

    public void setUpdatedAt(String updatedAt) {
        this.updatedAt = updatedAt;
    }

    public String getCreateUser() {
        return createUser;
    }

    public void setCreateUser(String createUser) {
        this.createUser = createUser;
    }
}

