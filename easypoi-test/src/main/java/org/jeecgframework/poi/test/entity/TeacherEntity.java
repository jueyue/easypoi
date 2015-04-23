package org.jeecgframework.poi.test.entity;

import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;

/**
 * @Title: Entity
 * @Description: 课程老师
 * @author JueYue
 * @date 2013-08-31 22:52:17
 * @version V1.0
 * 
 */
@SuppressWarnings("serial")
@ExcelTarget("teacherEntity")
public class TeacherEntity implements java.io.Serializable {
    /** id */
    @Excel(name = "老师ID_teacherEntity,老师属性_courseEntity", orderNum = "2", needMerge = false)
    private String id;
    /** name */
    @Excel(name = "老师姓名_yuwen,数学老师_shuxue", orderNum = "2", mergeVertical = true)
    private String name;

    /*
     * @Excel(exportName="老师照片",orderNum="3",exportType=2,exportFieldHeight=15,
     * exportFieldWidth=20) private String pic;
     */

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
     * @return: java.lang.String name
     */
    public String getName() {
        return this.name;
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
     * @param: java.lang.String name
     */
    public void setName(String name) {
        this.name = name;
    }

    /*
     * public String getPic() { // if(StringUtils.isEmpty(pic)){ // pic =
     * "plug-in/login/images/jeecg.png"; // } return pic; }
     * 
     * public void setPic(String pic) { this.pic = pic; }
     */
}
