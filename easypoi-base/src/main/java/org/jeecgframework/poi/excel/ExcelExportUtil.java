package org.jeecgframework.poi.excel;

import java.util.Collection;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.jeecgframework.poi.excel.export.ExcelExportServer;
import org.jeecgframework.poi.excel.export.template.ExcelExportOfTemplateUtil;

/**
 * excel 导出工具类
 * 
 * @author jueyue
 * @version 1.0
 * @date 2013-10-17
 */
public final class ExcelExportUtil {

	/**
	 * 一个excel 创建多个sheet
	 * 
	 * @param list
	 *            多个Map key title 对应表格Title key entity 对应表格对应实体 key data
	 *            Collection 数据
	 * @return
	 */
	public static HSSFWorkbook exportExcel(List<Map<String, Object>> list) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		ExcelExportServer server = new ExcelExportServer();
		for (Map<String, Object> map : list) {
			server.createSheet(workbook,
					(ExportParams) map.get("title"),
					(Class<?>) map.get("entity"),
					(Collection<?>) map.get("data"));
		}
		return workbook;
	}

	/**
	 * @param entity
	 *            表格标题属性
	 * @param pojoClass
	 *            Excel对象Class
	 * @param dataSet
	 *            Excel对象数据List
	 */
	public static HSSFWorkbook exportExcel(ExportParams entity,
			Class<?> pojoClass, Collection<?> dataSet) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		new ExcelExportServer().createSheet(workbook, entity, pojoClass, dataSet);
		return workbook;
	}

	/**
	 * 导出文件通过模板解析
	 * 
	 * @param params
	 *            导出参数类
	 * @param pojoClass
	 *            对应实体
	 * @param dataSet
	 *            实体集合
	 * @param map
	 *            模板集合
	 * @return
	 */
	public static Workbook exportExcel(TemplateExportParams params,
			Class<?> pojoClass, Collection<?> dataSet, Map<String, Object> map) {
		return new ExcelExportOfTemplateUtil().createExcleByTemplate(params,
				pojoClass, dataSet, map);
	}

	/**
	 * 导出文件通过模板解析只有模板,没有集合
	 * 
	 * @param params
	 *            导出参数类
	 * @param map
	 *            模板集合
	 * @return
	 */
	public static Workbook exportExcel(TemplateExportParams params,
			Map<String, Object> map) {
		return new ExcelExportOfTemplateUtil().createExcleByTemplate(params,
				null, null, map);
	}

	
}
