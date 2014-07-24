package org.jeecgframework.poi.excel.entity.params;

import java.util.List;
/**
 * excel 导入工具类,对cell类型做映射
 * @author jueyue
 * @version 1.0 2013年8月24日
 */
public class ExcelImportEntity extends ExcelBaseEntity {
	/**
	 * 对应 Collection NAME
	 */
	private String collectionName;
	/**
	 * 保存图片的地址
	 */
	private String  saveUrl;
	/**
	 * 保存图片的类型,1是文件,2是数据库
	 */
	private int  saveType;
	/**
	 * 对应exportType
	 */
	private String classType;
	/**
	 * 校驗參數
	 */
	private ExcelVerifyEntity verify;
	
	private List<ExcelImportEntity> list;

	public List<ExcelImportEntity> getList() {
		return list;
	}

	public void setList(List<ExcelImportEntity> list) {
		this.list = list;
	}

	public String getSaveUrl() {
		return saveUrl;
	}

	public void setSaveUrl(String saveUrl) {
		this.saveUrl = saveUrl;
	}

	public String getClassType() {
		return classType;
	}

	public void setClassType(String classType) {
		this.classType = classType;
	}

	public String getCollectionName() {
		return collectionName;
	}

	public void setCollectionName(String collectionName) {
		this.collectionName = collectionName;
	}

	public int getSaveType() {
		return saveType;
	}

	public void setSaveType(int saveType) {
		this.saveType = saveType;
	}

	public ExcelVerifyEntity getVerify() {
		return verify;
	}

	public void setVerify(ExcelVerifyEntity verify) {
		this.verify = verify;
	}

}
