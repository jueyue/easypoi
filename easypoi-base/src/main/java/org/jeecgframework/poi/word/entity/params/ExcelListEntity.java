package org.jeecgframework.poi.word.entity.params;

import java.util.List;

import org.jeecgframework.poi.excel.entity.ExcelBaseParams;
import org.jeecgframework.poi.handler.inter.IExcelDataHandler;

/**
 * Excel 导出对象
 * 
 * @author JueYue
 * @date 2014年8月9日 下午10:21:13
 */
public class ExcelListEntity extends ExcelBaseParams {

	public ExcelListEntity() {

	}

	public ExcelListEntity(List<?> list, Class<?> clazz) {
		this.list = list;
		this.clazz = clazz;
	}

	public ExcelListEntity(List<?> list, Class<?> clazz, int headRows) {
		this.list = list;
		this.clazz = clazz;
		this.headRows = headRows;
	}

	public ExcelListEntity(List<?> list, Class<?> clazz,
			IExcelDataHandler dataHanlder) {
		this.list = list;
		this.clazz = clazz;
		setDataHanlder(dataHanlder);
	}

	public ExcelListEntity(List<?> list, Class<?> clazz,
			IExcelDataHandler dataHanlder, int headRows) {
		this.list = list;
		this.clazz = clazz;
		this.headRows = headRows;
		setDataHanlder(dataHanlder);
	}

	/**
	 * 数据源
	 */
	private List<?> list;
	/**
	 * 实体类对象
	 */
	private Class<?> clazz;
	/**
	 * 表头行数
	 */
	private int headRows = 1;

	public List<?> getList() {
		return list;
	}

	public void setList(List<?> list) {
		this.list = list;
	}

	public Class<?> getClazz() {
		return clazz;
	}

	public void setClazz(Class<?> clazz) {
		this.clazz = clazz;
	}

	public int getHeadRows() {
		return headRows;
	}

	public void setHeadRows(int headRows) {
		this.headRows = headRows;
	}

}
