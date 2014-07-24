package org.jeecgframework.poi.excel.entity.params;

import java.util.List;

/**
 * excel 导出工具类,对cell类型做映射
 * 
 * @author jueyue
 * @version 1.0 2013年8月24日
 */
public class ExcelExportEntity extends ExcelBaseEntity {

	private int width;
	private int height;
	/**
	 * 图片的类型,1是文件,2是数据库
	 */
	private int exportImageType;
	/**
	 * 排序顺序
	 */
	private int orderNum;
	/**
	 * 是否支持换行
	 */
	private boolean isWrap;
	/**
	 * 是否需要合并
	 */
	private boolean needMerge;
	/**
	 * 单元格纵向合并
	 */
	private boolean mergeVertical;
	/**
	 * 合并依赖
	 */
	private int[] mergeRely;
	/**
	 * cell 函数
	 */
	private String cellFormula;

	private List<ExcelExportEntity> list;

	public int getWidth() {
		return width;
	}

	public void setWidth(int width) {
		this.width = width;
	}

	public List<ExcelExportEntity> getList() {
		return list;
	}

	public void setList(List<ExcelExportEntity> list) {
		this.list = list;
	}

	public int getHeight() {
		return height;
	}

	public void setHeight(int height) {
		this.height = height;
	}

	public int getOrderNum() {
		return orderNum;
	}

	public void setOrderNum(int orderNum) {
		this.orderNum = orderNum;
	}

	public boolean isWrap() {
		return isWrap;
	}

	public void setWrap(boolean isWrap) {
		this.isWrap = isWrap;
	}

	public boolean isNeedMerge() {
		return needMerge;
	}

	public void setNeedMerge(boolean needMerge) {
		this.needMerge = needMerge;
	}

	public int[] getMergeRely() {
		return mergeRely;
	}

	public void setMergeRely(int[] mergeRely) {
		this.mergeRely = mergeRely;
	}

	public int getExportImageType() {
		return exportImageType;
	}

	public void setExportImageType(int exportImageType) {
		this.exportImageType = exportImageType;
	}

	public String getCellFormula() {
		return cellFormula;
	}

	public void setCellFormula(String cellFormula) {
		this.cellFormula = cellFormula;
	}

	public boolean isMergeVertical() {
		return mergeVertical;
	}

	public void setMergeVertical(boolean mergeVertical) {
		this.mergeVertical = mergeVertical;
	}

}
