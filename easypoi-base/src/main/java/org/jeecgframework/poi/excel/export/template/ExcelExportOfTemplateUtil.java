package org.jeecgframework.poi.excel.export.template;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.cache.ExcelCache;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.export.base.ExcelExportBase;
import org.jeecgframework.poi.exception.excel.ExcelExportException;
import org.jeecgframework.poi.exception.excel.enums.ExcelExportEnum;
import org.jeecgframework.poi.util.POIPublicUtil;

/**
 * Excel 导出根据模板导出
 * 
 * @author JueYue
 * @date 2013-10-17
 * @version 1.0
 */
public final class ExcelExportOfTemplateUtil extends ExcelExportBase {
	

	public Workbook createExcleByTemplate(TemplateExportParams params,
			Class<?> pojoClass, Collection<?> dataSet, Map<String, Object> map) {
		// step 1. 判断模板的地址
		if (params == null || map == null
				|| StringUtils.isEmpty(params.getTemplateUrl())) {
			throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
		}
		Workbook wb = null;
		// step 2. 判断模板的Excel类型,解析模板
		try {
			wb = getCloneWorkBook(params);
			if (StringUtils.isNotEmpty(params.getSheetName())) {
				wb.setSheetName(0, params.getSheetName());
			}
			// step 3. 解析模板
			parseTemplate(wb.getSheetAt(0), map);
			if (dataSet != null) {
				// step 4. 正常的数据填充
				dataHanlder = params.getDataHanlder();
				if (dataHanlder != null) {
					needHanlderList = Arrays.asList(dataHanlder
							.getNeedHandlerFields());
				}
				addDataToSheet(params, pojoClass, dataSet, wb.getSheetAt(0), wb);
			}
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
		return wb;
	}

	/**
	 * 克隆excel防止操作原对象,workbook无法克隆,只能对excel进行克隆
	 * 
	 * @param params
	 * @throws Exception
	 * @Author JueYue
	 * @date 2013-11-11
	 */
	private Workbook getCloneWorkBook(TemplateExportParams params)
			throws Exception {
		return ExcelCache.getWorkbook(params.getTemplateUrl(),
				params.getSheetNum());

	}

	/**
	 * 往Sheet 填充正常数据,根据表头信息 使用导入的部分逻辑,坐对象映射
	 * 
	 * @param params
	 * @param pojoClass
	 * @param dataSet
	 * @param workbook
	 */
	private void addDataToSheet(TemplateExportParams params,
			Class<?> pojoClass, Collection<?> dataSet, Sheet sheet,
			Workbook workbook) throws Exception {
		// 获取表头数据
		Map<String, Integer> titlemap = getTitleMap(params, sheet);
		Drawing patriarch = sheet.createDrawingPatriarch();
		// 得到所有字段
		Field fileds[] = POIPublicUtil.getClassFields(pojoClass);
		ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
		String targetId = null;
		if (etarget != null) {
			targetId = etarget.value();
		}
		// 获取实体对象的导出数据
		List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
		getAllExcelField(null, targetId, fileds, excelParams, pojoClass, null);
		// 根据表头进行筛选排序
		sortAndFilterExportField(excelParams, titlemap);
		short rowHeight = getRowHeight(excelParams);
		Iterator<?> its = dataSet.iterator();
		int index = sheet.getLastRowNum() + 1, titleHeight = index;
		while (its.hasNext()) {
			Object t = its.next();
			index += createCells(patriarch, index, t, excelParams, sheet,
					workbook, null, rowHeight);
		}
		// 合并同类项
		mergeCells(sheet, excelParams, titleHeight);
	}

	/**
	 * 对导出序列进行排序和塞选
	 * 
	 * @param excelParams
	 * @param titlemap
	 * @return
	 */
	private void sortAndFilterExportField(List<ExcelExportEntity> excelParams,
			Map<String, Integer> titlemap) {
		for (int i = excelParams.size() - 1; i >= 0; i--) {
			if (excelParams.get(i).getList() != null
					&& excelParams.get(i).getList().size() > 0) {
				sortAndFilterExportField(excelParams.get(i).getList(), titlemap);
				if (excelParams.get(i).getList().size() == 0) {
					excelParams.remove(i);
				} else {
					excelParams.get(i).setOrderNum(i);
				}
			} else {
				if (titlemap.containsKey(excelParams.get(i).getName())) {
					excelParams.get(i).setOrderNum(i);
				} else {
					excelParams.remove(i);
				}
			}
		}
		sortAllParams(excelParams);
	}

	/**
	 * 获取表头数据,设置表头的序号
	 * 
	 * @param params
	 * @param sheet
	 * @return
	 */
	private Map<String, Integer> getTitleMap(TemplateExportParams params,
			Sheet sheet) {
		Row row = null;
		Iterator<Cell> cellTitle;
		Map<String, Integer> titlemap = new HashMap<String, Integer>();
		for (int j = 0; j < params.getHeadingRows(); j++) {
			row = sheet.getRow(j + params.getHeadingStartRow());
			cellTitle = row.cellIterator();
			int i = row.getFirstCellNum();
			while (cellTitle.hasNext()) {
				Cell cell = cellTitle.next();
				String value = cell.getStringCellValue();
				if (!StringUtils.isEmpty(value)) {
					titlemap.put(value, i);
				}
				i = i + 1;
			}
		}
		return titlemap;

	}

	private void parseTemplate(Sheet sheet, Map<String, Object> map)
			throws Exception {
		Iterator<Row> rows = sheet.rowIterator();
		Row row;
		while (rows.hasNext()) {
			row = rows.next();
			for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
				setValueForCellByMap(row.getCell(i), map);
			}
		}
	}

	/**
	 * 给每个Cell通过解析方式set值
	 * 
	 * @param cell
	 * @param map
	 */
	private void setValueForCellByMap(Cell cell, Map<String, Object> map)
			throws Exception {
		String oldString;
		try {// step 1. 判断这个cell里面是不是函数
			oldString = cell.getStringCellValue();
		} catch (Exception e) {
			return;
		}
		if (oldString != null && oldString.indexOf("{{") != -1) {
			// setp 2. 判断是否含有解析函数
			String params;
			while (oldString.indexOf("{{") != -1) {
				params = oldString.substring(oldString.indexOf("{{") + 2,
						oldString.indexOf("}}"));
				oldString = oldString.replace("{{" + params + "}}",
						getParamsValue(params.trim(), map));
			}
			cell.setCellValue(oldString);
		}
	}

	/**
	 * 获取参数值
	 * 
	 * @param params
	 * @param map
	 * @return
	 */
	private String getParamsValue(String params, Map<String, Object> map)
			throws Exception {
		if (params.indexOf(".") != -1) {
			String[] paramsArr = params.split("\\.");
			return getValueDoWhile(map.get(paramsArr[0]), paramsArr, 1);
		}
		return map.containsKey(params) ? map.get(params).toString() : "";
	}

	/**
	 * 通过遍历过去对象值
	 * 
	 * @param object
	 * @param paramsArr
	 * @param index
	 * @return
	 * @throws Exception
	 * @throws java.lang.reflect.InvocationTargetException
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 */
	@SuppressWarnings("rawtypes")
	private String getValueDoWhile(Object object, String[] paramsArr, int index)
			throws Exception {
		if (object == null) {
			return "";
		}
		if (object instanceof Map) {
			object = ((Map) object).get(paramsArr[index]);
		} else {
			object = POIPublicUtil.getMethod(paramsArr[index],
					object.getClass()).invoke(object, new Object[] {});
		}
		return (index == paramsArr.length - 1) ? (object == null ? "" : object
				.toString()) : getValueDoWhile(object, paramsArr, ++index);
	}

}
