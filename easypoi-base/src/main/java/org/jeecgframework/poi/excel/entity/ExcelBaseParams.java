package org.jeecgframework.poi.excel.entity;

import org.jeecgframework.poi.handler.inter.IExcelDataHandler;

/**
 * 基础参数
 * @author JueYue
 * @date 2014年6月20日 下午1:56:52
 */
public class ExcelBaseParams {

    /**
     * 03版本Excel
     */
    public static final String      HSSF = "HSSF";
    /**
     * 07版本Excel
     */
    public static final String      XSSF = "XSSF";
    /**
     * 数据处理接口,以此为主,replace,format都在这后面
     */
    private IExcelDataHandler dataHanlder;

    public IExcelDataHandler getDataHanlder() {
        return dataHanlder;
    }

    public void setDataHanlder(IExcelDataHandler dataHanlder) {
        this.dataHanlder = dataHanlder;
    }

}
