package com.zhangyibo.formatreader;

import com.alibaba.fastjson.JSONObject;
import com.zhangyibo.formatreader.format.excel.ExcelReader;

import java.io.File;
import java.io.IOException;

public class FormatReader {
    public static void main(String[] args) throws IOException {
        final String filePath = "./sample/src/main/resources/xiaoxuesheng.xlsx";
        ExcelReader<Student> reader  = new ExcelReader<>();
        reader.open(new File(filePath), Student.class);
        Student student;
        System.out.println("=============== BEGIN READ FILE ================");
        while ((student = reader.nextObject()) != null) {
            String jsonString = JSONObject.toJSONString(student);
            System.out.println("Read xml , line :  " + reader.index() + " : " +  jsonString);
        }
        System.out.println("===============  END READ FILE  ================");
    }
}
