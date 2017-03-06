/**
 * 
 */
package org.jeecgframework.poi.excel.entity;

import java.util.List;
import java.util.Map;

import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;

/**
 * @author xfworld
 * @since 2016-5-26
 * @version 1.0
 */
public class ExportExcelItem {
	private List<ExcelExportEntity> entityList;
	private List<Map<String, Object>> resultList;

	public List<ExcelExportEntity> getEntityList() {
		return this.entityList;
	}

	public void setEntityList(List<ExcelExportEntity> entityList) {
		this.entityList = entityList;
	}

	public List<Map<String, Object>> getResultList() {
		return this.resultList;
	}

	public void setResultList(List<Map<String, Object>> resultList) {
		this.resultList = resultList;
	}
}
