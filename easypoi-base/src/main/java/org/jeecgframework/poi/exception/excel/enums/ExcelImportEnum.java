package org.jeecgframework.poi.exception.excel.enums;

/**
 * 导出异常类型枚举
 * @author JueYue
 * @date 2014年6月19日 下午10:59:51
 */
public enum ExcelImportEnum {

    GET_VALUE_ERROR("Excel 值获取失败");

    private String msg;

    ExcelImportEnum(String msg) {
        this.msg = msg;
    }

    public String getMsg() {
        return msg;
    }

    public void setMsg(String msg) {
        this.msg = msg;
    }

}
