package org.jeecgframework.poi.exception.excel;

import org.jeecgframework.poi.exception.excel.enums.ExcelImportEnum;

/**
 * 导入异常
 * @author JueYue
 * @date 2014年6月29日 下午2:23:43
 */
public class ExcelImportException extends RuntimeException {

    private static final long serialVersionUID = 1L;

    private ExcelImportEnum   type;

    public ExcelImportException() {
        super();
    }

    public ExcelImportException(ExcelImportEnum type) {
        super(type.getMsg());
        this.type = type;
    }

    public ExcelImportException(ExcelImportEnum type, Throwable cause) {
        super(type.getMsg(), cause);
    }

    public ExcelImportException(String message) {
        super(message);
    }

    public ExcelImportException(String message, ExcelImportEnum type) {
        super(message);
        this.type = type;
    }

    public ExcelImportEnum getType() {
        return type;
    }

    public void setType(ExcelImportEnum type) {
        this.type = type;
    }

}
