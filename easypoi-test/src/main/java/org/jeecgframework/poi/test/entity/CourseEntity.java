package org.jeecgframework.poi.test.entity;

import java.util.List;

import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.annotation.ExcelCollection;
import org.jeecgframework.poi.excel.annotation.ExcelEntity;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.annotation.ExcelVerify;

/**
 * @Title: Entity
 * @Description: 课程
 * @author JueYue
 * @date 2013-08-31 22:53:07
 * @version V1.0
 * 
 */
@SuppressWarnings("serial")
@ExcelTarget("courseEntity")
public class CourseEntity implements java.io.Serializable {
    /** 主键 */
    private String              id;
    /** 课程名称 */
    @Excel(name = "课程名称", orderNum = "1", width = 15, needMerge = true)
    private String              name;
    /** 老师主键 */
    @ExcelEntity(id = "yuwen")
    @ExcelVerify()
    private TeacherEntity       teacher;
    /** 老师主键 */
    //@ExcelEntity(id = "shuxue")
    private TeacherEntity       shuxueteacher;

    @ExcelCollection(name = "", orderNum = "4")
    private List<StudentEntity> students;

    /**
     * 方法: 取得java.lang.String
     * 
     * @return: java.lang.String 主键
     */

    public String getId() {
        return this.id;
    }

    /**
     * 方法: 取得java.lang.String
     * 
     * @return: java.lang.String 课程名称
     */
    public String getName() {
        return this.name;
    }

    public TeacherEntity getShuxueteacher() {
        return shuxueteacher;
    }

    public List<StudentEntity> getStudents() {
        return students;
    }

    /**
     * 方法: 取得java.lang.String
     * 
     * @return: java.lang.String 老师主键
     */
    public TeacherEntity getTeacher() {
        return teacher;
    }

    /**
     * 方法: 设置java.lang.String
     * 
     * @param: java.lang.String 主键
     */
    public void setId(String id) {
        this.id = id;
    }

    /**
     * 方法: 设置java.lang.String
     * 
     * @param: java.lang.String 课程名称
     */
    public void setName(String name) {
        this.name = name;
    }

    public void setShuxueteacher(TeacherEntity shuxueteacher) {
        this.shuxueteacher = shuxueteacher;
    }

    public void setStudents(List<StudentEntity> students) {
        this.students = students;
    }

    /**
     * 方法: 设置java.lang.String
     * 
     * @param: java.lang.String 老师主键
     */
    public void setTeacher(TeacherEntity teacher) {
        this.teacher = teacher;
    }
}
