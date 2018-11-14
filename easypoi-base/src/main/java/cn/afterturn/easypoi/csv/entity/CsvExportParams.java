package cn.afterturn.easypoi.csv.entity;

import cn.afterturn.easypoi.excel.entity.ExcelBaseParams;

/**
 * CSV 导入参数
 *
 * @author by jueyue on 18-10-3.
 */
public class CsvExportParams extends ExcelBaseParams {

    public static final String UTF8 = "utf-8";
    public static final String GBK = "gbk";
    public static final String GB2312 = "gb2312";

    private String encoding = UTF8;

    /**
     * 分隔符
     */
    private String spiltMark = ",";

    /**
     * 字符串标识符
     */
    private String textMark = "\"";

    /**
     * 表格标题行数,默认0
     */
    private int titleRows = 0;

    /**
     * 表头行数,默认1
     */
    private int headRows = 1;
    /**
     * 过滤的属性
     */
    private String[] exclusions;

    /**
     * 是否创建表头
     */
    private boolean isCreateHeadRows = true;

    public CsvExportParams() {

    }

    public CsvExportParams(String encoding) {
        this.encoding = encoding;
    }

    public String getEncoding() {
        return encoding;
    }

    public void setEncoding(String encoding) {
        this.encoding = encoding;
    }

    public String getSpiltMark() {
        return spiltMark;
    }

    public void setSpiltMark(String spiltMark) {
        this.spiltMark = spiltMark;
    }

    public String getTextMark() {
        return textMark;
    }

    public void setTextMark(String textMark) {
        this.textMark = textMark;
    }

    public int getTitleRows() {
        return titleRows;
    }

    public void setTitleRows(int titleRows) {
        this.titleRows = titleRows;
    }

    public int getHeadRows() {
        return headRows;
    }

    public void setHeadRows(int headRows) {
        this.headRows = headRows;
    }

    public String[] getExclusions() {
        return exclusions;
    }

    public void setExclusions(String[] exclusions) {
        this.exclusions = exclusions;
    }

    public boolean isCreateHeadRows() {
        return isCreateHeadRows;
    }

    public void setCreateHeadRows(boolean createHeadRows) {
        isCreateHeadRows = createHeadRows;
    }
}
