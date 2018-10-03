package cn.afterturn.easypoi.csv.entity;

import cn.afterturn.easypoi.excel.entity.ExcelBaseParams;
import cn.afterturn.easypoi.handler.inter.IExcelVerifyHandler;

/**
 * CSV 导入参数
 *
 * @author by jueyue on 18-10-3.
 */
public class CsvImportParams extends ExcelBaseParams {

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
     * 字段真正值和列标题之间的距离 默认0
     */
    private int startRows = 0;
    /**
     * 校验组
     */
    private Class[] verifyGroup = null;
    /**
     * 是否需要校验上传的Excel,默认为false
     */
    private boolean needVerify = false;
    /**
     * 校验处理接口
     */
    private IExcelVerifyHandler verifyHandler;
    /**
     * 最后的无效行数
     */
    private int lastOfInvalidRow = 0;

    /**
     * 主键设置,如何这个cell没有值,就跳过 或者认为这个是list的下面的值
     * 大家不理解，去掉这个
     */

    private Integer keyIndex = null;

    public CsvImportParams() {

    }

    public CsvImportParams(String encoding) {
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

    public int getStartRows() {
        return startRows;
    }

    public void setStartRows(int startRows) {
        this.startRows = startRows;
    }

    public Class[] getVerifyGroup() {
        return verifyGroup;
    }

    public void setVerifyGroup(Class[] verifyGroup) {
        this.verifyGroup = verifyGroup;
    }

    public boolean isNeedVerify() {
        return needVerify;
    }

    public void setNeedVerify(boolean needVerify) {
        this.needVerify = needVerify;
    }

    public IExcelVerifyHandler getVerifyHandler() {
        return verifyHandler;
    }

    public void setVerifyHandler(IExcelVerifyHandler verifyHandler) {
        this.verifyHandler = verifyHandler;
    }

    public int getLastOfInvalidRow() {
        return lastOfInvalidRow;
    }

    public void setLastOfInvalidRow(int lastOfInvalidRow) {
        this.lastOfInvalidRow = lastOfInvalidRow;
    }

    public Integer getKeyIndex() {
        return keyIndex;
    }

    public void setKeyIndex(Integer keyIndex) {
        this.keyIndex = keyIndex;
    }
}
