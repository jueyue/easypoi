package org.jeecgframework.poi.entity;

import java.util.Date;

import org.jeecgframework.poi.excel.annotation.Excel;

/**
 * @author jueyue
 * @version V1.0
 * @Title: Entity
 * @Description: 学生
 * @date 2013-08-31 22:53:34
 */
@SuppressWarnings("serial")
public class StudentEntity implements java.io.Serializable {
    /**
     * id
     */
    private String       id;
    /**
     * 学生姓名
     */
    @Excel(name = "学生姓名", height = 20, width = 30)
    private String       name;
    /**
     * 学生性别
     */
    @Excel(name = "学生性别", replace = { "男_1", "女_2" })
    private int          sex;

    @Excel(name = "出生日期", format = "yyyy-MM-dd HH:mm:ss", mergeVertical = true)
    private Date         birthday;
    /**
     * 课程主键
     */
    private CourseEntity course;

    public Date getBirthday() {
        return birthday;
    }

    public CourseEntity getCourse() {
        return course;
    }

    /**
     * 方法: 取得java.lang.String
     * 
     * @return: java.lang.String id
     */
    public String getId() {
        return this.id;
    }

    /**
     * 方法: 取得java.lang.String
     * 
     * @return: java.lang.String 学生姓名
     */
    public String getName() {
        return this.name;
    }

    /**
     * 方法: 取得java.lang.String
     * 
     * @return: java.lang.String 学生性别
     */
    public int getSex() {
        return this.sex;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    public void setCourse(CourseEntity course) {
        this.course = course;
    }

    /**
     * 方法: 设置java.lang.String
     * 
     * @param: java.lang.String id
     */
    public void setId(String id) {
        this.id = id;
    }

    /**
     * 方法: 设置java.lang.String
     * 
     * @param: java.lang.String 学生姓名
     */
    public void setName(String name) {
        this.name = name;
    }

    /**
     * 方法: 设置java.lang.String
     * 
     * @param: java.lang.String 学生性别
     */
    public void setSex(int sex) {
        this.sex = sex;
    }
}
