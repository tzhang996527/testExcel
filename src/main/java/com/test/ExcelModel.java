package com.test;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

import java.util.Date;

public class ExcelModel extends BaseRowModel{
    @ExcelProperty(value = "姓名", index = 0)//0 代表第一列
    private String name;

    @ExcelProperty(value = "昵称", index = 1)
    private String nickName;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getNickName() {
        return nickName;
    }

    public void setNickName(String nickName) {
        this.nickName = nickName;
    }

//    @ExcelProperty(value = "密码", index = 2)
//    private String password;
//
//    //不支持LocalDate
//    @ExcelProperty(value = "生日", index = 3, format = "yyyy/MM/dd")
//    private Date birthday;

}
