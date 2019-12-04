package com.zhangyibo.formatreader;

import com.zhangyibo.formatreader.base.FormatClass;
import com.zhangyibo.formatreader.base.FormatField;

@FormatClass
public class Student {
    @FormatField("姓名")
    public String name;
    @FormatField("年龄")
    public int age;
    @FormatField("学号")
    public String studentNo;
    @FormatField("性别")
    public String sex;
    @FormatField("年级")
    public String grade;
    @FormatField("语文")
    public double langScore;
    @FormatField("数学")
    public double mathScore;
    @FormatField("英语")
    public double englishScore;
}
