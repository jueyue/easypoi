package org.jeecgframework.poi.exception.excel;

import org.jeecgframework.poi.exception.excel.enums.ExcelExportEnum;

/**
 * 导出异常
 * 
 * @author JueYue
 * @date 2014年6月19日 下午10:56:18
 */
public class ExcelExportException extends RuntimeException {

	private static final long serialVersionUID = 1L;

	private ExcelExportEnum type;

	public ExcelExportException() {
		super();
	}

	public ExcelExportException(String message) {
		super(message);
	}

	public ExcelExportException(ExcelExportEnum type) {
		super(type.getMsg());
		this.type = type;
	}

	public ExcelExportException(String message, ExcelExportEnum type) {
		super(message);
		this.type = type;
	}
	
	public ExcelExportException(ExcelExportEnum type, Throwable cause) {
        super(type.getMsg(), cause);
    }

	public ExcelExportEnum getType() {
		return type;
	}

	public void setType(ExcelExportEnum type) {
		this.type = type;
	}

}
