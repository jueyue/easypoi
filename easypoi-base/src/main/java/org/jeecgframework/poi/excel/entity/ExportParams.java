package org.jeecgframework.poi.excel.entity;

import org.apache.poi.hssf.util.HSSFColor;
import org.jeecgframework.poi.excel.entity.vo.PoiBaseConstants;

/**
 * Excel 导出参数
 * 
 * @author jueyue
 * @version 1.0 2013年8月24日
 */
public class ExportParams extends ExcelBaseParams {

    /**
     * 表格名称
     */
    private String   title;

    /**
     * 表格名称
     */
    private short    titleHeight       = 20;

    /**
     * 第二行名称
     */
    private String   secondTitle;

    /**
     * 表格名称
     */
    private short    secondTitleHeight = 8;
    /**
     * sheetName
     */
    private String   sheetName;
    /**
     * 过滤的属性
     */
    private String[] exclusions;
    /**
     * 是否添加需要需要
     */
    private boolean  addIndex;
    /**
     * 表头颜色
     */
    private short    color             = HSSFColor.WHITE.index;
    /**
     * 属性说明行的颜色 例如:HSSFColor.SKY_BLUE.index 默认
     */
    private short    headerColor       = HSSFColor.SKY_BLUE.index;
    /**
     * Excel 导出版本
     */
    private String   type              = PoiBaseConstants.HSSF;

    public ExportParams() {

    }

    public ExportParams(String title, String sheetName) {
        this.title = title;
        this.sheetName = sheetName;
    }

    public ExportParams(String title, String secondTitle, String sheetName) {
        this.title = title;
        this.secondTitle = secondTitle;
        this.sheetName = sheetName;
    }

    public short getColor() {
        return color;
    }

    public String[] getExclusions() {
        return exclusions;
    }

    public short getHeaderColor() {
        return headerColor;
    }

    public String getSecondTitle() {
        return secondTitle;
    }

    public short getSecondTitleHeight() {
        return (short) (secondTitleHeight * 50);
    }

    public String getSheetName() {
        return sheetName;
    }

    public String getTitle() {
        return title;
    }

    public short getTitleHeight() {
        return (short) (titleHeight * 50);
    }

    public boolean isAddIndex() {
        return addIndex;
    }

    public void setAddIndex(boolean addIndex) {
        this.addIndex = addIndex;
    }

    public void setColor(short color) {
        this.color = color;
    }

    public void setExclusions(String[] exclusions) {
        this.exclusions = exclusions;
    }

    public void setHeaderColor(short headerColor) {
        this.headerColor = headerColor;
    }

    public void setSecondTitle(String secondTitle) {
        this.secondTitle = secondTitle;
    }

    public void setSecondTitleHeight(short secondTitleHeight) {
        this.secondTitleHeight = secondTitleHeight;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public void setTitleHeight(short titleHeight) {
        this.titleHeight = titleHeight;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }
}
