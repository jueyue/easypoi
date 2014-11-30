package org.jeecgframework.poi.excel.entity;

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
    private String templateUrl;

    /**
     * 需要导出的第几个 sheetNum,默认是第0个
     */
    private int    sheetNum        = 0;

    /**
     * 这只sheetName 不填就使用原来的
     */
    private String sheetName;

    /**
     * 表格列标题行数,默认1
     */
    private int    headingRows     = 1;

    /**
     * 表格列标题开始行,默认1
     */
    private int    headingStartRow = 2;

    public TemplateExportParams() {

    }

    public TemplateExportParams(String templateUrl) {
        this.templateUrl = templateUrl;
    }
    public TemplateExportParams(String templateUrl, int sheetNum) {
        this.templateUrl = templateUrl;
        this.sheetNum = sheetNum;
    }
    public TemplateExportParams(String templateUrl, String sheetName) {
        this.templateUrl = templateUrl;
        this.sheetName = sheetName;
    }
    public TemplateExportParams(String templateUrl, String sheetName, int sheetNum) {
        this.templateUrl = templateUrl;
        this.sheetName = sheetName;
        this.sheetNum = sheetNum;
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

    public int getSheetNum() {
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

    public void setSheetNum(int sheetNum) {
        this.sheetNum = sheetNum;
    }

    public void setTemplateUrl(String templateUrl) {
        this.templateUrl = templateUrl;
    }

}
