package org.jeecgframework.poi.excel.entity;

import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerDefaultImpl;

/**
 * 模板导出参数设置
 * 
 * @author JueYue
 * @date 2013-10-17
 * @version 1.0
 */
public class TemplateExportParams extends ExcelBaseParams {

    /**
     * 模板的路径
     */
    private String    templateUrl;

    /**
     * 需要导出的第几个 sheetNum,默认是第0个
     */
    private Integer[] sheetNum        = new Integer[] { 0 };

    /**
     * 这只sheetName 不填就使用原来的
     */
    private String    sheetName;

    /**
     * 表格列标题行数,默认1
     */
    private int       headingRows     = 1;

    /**
     * 表格列标题开始行,默认1
     */
    private int       headingStartRow = 1;
    /**
     * Excel 导出style
     */
    private Class<?>  style           = ExcelExportStylerDefaultImpl.class;

    public TemplateExportParams() {

    }

    public TemplateExportParams(String templateUrl, Integer... sheetNum) {
        this.templateUrl = templateUrl;
        if (sheetNum != null && sheetNum.length > 0) {
            this.sheetNum = sheetNum;
        }
    }

    public TemplateExportParams(String templateUrl, String sheetName, Integer... sheetNum) {
        this.templateUrl = templateUrl;
        this.sheetName = sheetName;
        if (sheetNum != null && sheetNum.length > 0) {
            this.sheetNum = sheetNum;
        }
    }

    public int getHeadingRows() {
        return headingRows;
    }

    public int getHeadingStartRow() {
        return headingStartRow;
    }

    public String getSheetName() {
        return sheetName;
    }

    public Integer[] getSheetNum() {
        return sheetNum;
    }

    public String getTemplateUrl() {
        return templateUrl;
    }

    public void setHeadingRows(int headingRows) {
        this.headingRows = headingRows;
    }

    public void setHeadingStartRow(int headingStartRow) {
        this.headingStartRow = headingStartRow;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public void setSheetNum(Integer[] sheetNum) {
        this.sheetNum = sheetNum;
    }

    public void setSheetNum(Integer sheetNum) {
        this.sheetNum = new Integer[] { sheetNum };
    }

    public void setTemplateUrl(String templateUrl) {
        this.templateUrl = templateUrl;
    }

    public Class<?> getStyle() {
        return style;
    }

    public void setStyle(Class<?> style) {
        this.style = style;
    }

}
