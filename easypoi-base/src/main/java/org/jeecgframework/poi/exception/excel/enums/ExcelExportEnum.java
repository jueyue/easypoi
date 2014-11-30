package org.jeecgframework.poi.exception.excel.enums;

/**
 * 导出异常类型枚举
 * @author JueYue
 * @date 2014年6月19日 下午10:59:51
 */
public enum ExcelExportEnum {

    PARAMETER_ERROR("Excel 导出   参数错误"), EXPORT_ERROR("Excel导出错误");

    private String msg;

    ExcelExportEnum(String msg) {
        this.msg = msg;
    }

    public String getMsg() {
        return msg;
    }

    public void setMsg(String msg) {
        this.msg = msg;
    }

}
