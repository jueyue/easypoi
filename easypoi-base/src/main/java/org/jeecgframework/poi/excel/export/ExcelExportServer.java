package org.jeecgframework.poi.excel.export;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.entity.vo.PoiBaseConstants;
import org.jeecgframework.poi.excel.export.base.ExcelExportBase;
import org.jeecgframework.poi.exception.excel.ExcelExportException;
import org.jeecgframework.poi.exception.excel.enums.ExcelExportEnum;
import org.jeecgframework.poi.util.POIPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Excel导出服务
 * 
 * @author JueYue
 * @date 2014年6月17日 下午5:30:54
 */
public class ExcelExportServer extends ExcelExportBase {

	private final static Logger logger = LoggerFactory
			.getLogger(ExcelExportServer.class);

	private static final short cellFormat = HSSFDataFormat
			.getBuiltinFormat("TEXT");

	// 最大行数,超过自动多Sheet
	private final int MAX_NUM = 60000;

	public void createSheet(HSSFWorkbook workbook, ExportParams entity,
			Class<?> pojoClass, Collection<?> dataSet) {
		if (logger.isDebugEnabled()) {
			logger.debug("Excel export start ,class is {}", pojoClass);
		}
		if (workbook == null || entity == null || pojoClass == null
				|| dataSet == null) {
			throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
		}
		Sheet sheet = null;
		try {
			sheet = workbook.createSheet(entity.getSheetName());
		} catch (Exception e) {
			// 重复遍历,出现了重名现象,创建非指定的名称Sheet
			sheet = workbook.createSheet();
		}
		try {
			dataHanlder = entity.getDataHanlder();
			if (dataHanlder != null) {
				needHanlderList = Arrays.asList(dataHanlder
						.getNeedHandlerFields());
			}
			// 创建表格属性
			Map<String, HSSFCellStyle> styles = createStyles(workbook);
			Drawing patriarch = sheet.createDrawingPatriarch();
			List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
			if (entity.isAddIndex()) {
				excelParams.add(indexExcelEntity());
			}
			// 得到所有字段
			Field fileds[] = POIPublicUtil.getClassFields(pojoClass);
			ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
			String targetId = etarget == null ? null : etarget.value();
			getAllExcelField(entity.getExclusions(), targetId, fileds,
					excelParams, pojoClass, null);
			sortAllParams(excelParams);
			int index = createHeaderAndTitle(entity, sheet, workbook,
					excelParams);
			int titleHeight = index;
			setCellWith(excelParams, sheet);
			short rowHeight = getRowHeight(excelParams);
			setCurrentIndex(1);
			Iterator<?> its = dataSet.iterator();
			List<Object> tempList = new ArrayList<Object>();
			while (its.hasNext()) {
				Object t = its.next();
				index += createCells(patriarch, index, t, excelParams, sheet,
						workbook, styles, rowHeight);
				tempList.add(t);
				if (index >= MAX_NUM)
					break;
			}
			mergeCells(sheet, excelParams, titleHeight);

			its = dataSet.iterator();
			for (int i = 0, le = tempList.size(); i < le; i++) {
				its.next();
				its.remove();
			}
			// 发现还有剩余list 继续循环创建Sheet
			if (dataSet.size() > 0) {
				createSheet(workbook, entity, pojoClass, dataSet);
			}

		} catch (Exception e) {
			logger.error(e.getMessage(),e.fillInStackTrace());
			throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR,
					e.getCause());
		}
	}

	private ExcelExportEntity indexExcelEntity() {
		ExcelExportEntity entity = new ExcelExportEntity();
		entity.setOrderNum(0);
		entity.setName("序号");
		entity.setWidth(10);
		entity.setFormat(PoiBaseConstants.IS_ADD_INDEX);
		return entity;
	}

	private int createHeaderAndTitle(ExportParams entity, Sheet sheet,
			HSSFWorkbook workbook, List<ExcelExportEntity> excelParams) {
		int rows = 0, feildWidth = getFieldWidth(excelParams);
		if (entity.getTitle() != null) {
			rows += createHeaderRow(entity, sheet, workbook, feildWidth);
		}
		rows += createTitleRow(entity, sheet, workbook, rows, excelParams);
		sheet.createFreezePane(0, rows, 0, rows);
		return rows;
	}

	/**
	 * 创建表头
	 * 
	 * @param title
	 * @param index
	 */
	private int createTitleRow(ExportParams title, Sheet sheet,
			HSSFWorkbook workbook, int index,
			List<ExcelExportEntity> excelParams) {
		Row row = sheet.createRow(index);
		int rows = getRowNums(excelParams);
		row.setHeight((short) 450);
		Row listRow = null;
		if (rows == 2) {
			listRow = sheet.createRow(index + 1);
			listRow.setHeight((short) 450);
		}
		int cellIndex = 0;
		CellStyle titleStyle = getTitleStyle(workbook, title);
		for (int i = 0, exportFieldTitleSize = excelParams.size(); i < exportFieldTitleSize; i++) {
			ExcelExportEntity entity = excelParams.get(i);
			createStringCell(row, cellIndex, entity.getName(), titleStyle,
					entity);
			if (entity.getList() != null) {
				List<ExcelExportEntity> sTitel = entity.getList();
				sheet.addMergedRegion(new CellRangeAddress(index, index,
						cellIndex, cellIndex + sTitel.size() - 1));
				for (int j = 0, size = sTitel.size(); j < size; j++) {
					createStringCell(listRow, cellIndex, sTitel.get(j)
							.getName(), titleStyle, entity);
					cellIndex++;
				}
			} else if (rows == 2) {
				sheet.addMergedRegion(new CellRangeAddress(index, index + 1,
						cellIndex, cellIndex));
			}
			cellIndex++;
		}
		return rows;

	}

	/**
	 * 判断表头是只有一行还是两行
	 * 
	 * @param excelParams
	 * @return
	 */
	private int getRowNums(List<ExcelExportEntity> excelParams) {
		for (int i = 0; i < excelParams.size(); i++) {
			if (excelParams.get(i).getList() != null) {
				return 2;
			}
		}
		return 1;
	}

	/**
	 * 创建 表头改变
	 * 
	 * @param entity
	 * @param sheet
	 * @param workbook
	 * @param feildWidth
	 */
	public int createHeaderRow(ExportParams entity, Sheet sheet,
			HSSFWorkbook workbook, int feildWidth) {
		Row row = sheet.createRow(0);
		row.setHeight(entity.getTitleHeight());
		createStringCell(row, 0, entity.getTitle(),
				getHeaderStyle(workbook, entity), null);
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, feildWidth));
		if (entity.getSecondTitle() != null) {
			row = sheet.createRow(1);
			row.setHeight(entity.getSecondTitleHeight());
			HSSFCellStyle style = workbook.createCellStyle();
			style.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			createStringCell(row, 0, entity.getSecondTitle(), style, null);
			sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, feildWidth));
			return 2;
		}
		return 1;
	}

	/**
	 * 字段说明的Style
	 * 
	 * @param workbook
	 * @return
	 */
	public HSSFCellStyle getTitleStyle(HSSFWorkbook workbook,
			ExportParams entity) {
		HSSFCellStyle titleStyle = workbook.createCellStyle();
		titleStyle.setFillForegroundColor(entity.getHeaderColor()); // 填充的背景颜色
		titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		titleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		titleStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND); // 填充图案
		titleStyle.setWrapText(true);
		return titleStyle;
	}

	/**
	 * 表明的Style
	 * 
	 * @param workbook
	 * @return
	 */
	public HSSFCellStyle getHeaderStyle(HSSFWorkbook workbook,
			ExportParams entity) {
		HSSFCellStyle titleStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontHeightInPoints((short) 24);
		titleStyle.setFont(font);
		titleStyle.setFillForegroundColor(entity.getColor());
		titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		titleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		return titleStyle;
	}

	public HSSFCellStyle getTwoStyle(HSSFWorkbook workbook, boolean isWarp) {
		HSSFCellStyle style = workbook.createCellStyle();
		style.setBorderLeft((short) 1); // 左边框
		style.setBorderRight((short) 1); // 右边框
		style.setBorderBottom((short) 1);
		style.setBorderTop((short) 1);
		style.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index); // 填充的背景颜色
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND); // 填充图案
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style.setDataFormat(cellFormat);
		if (isWarp) {
			style.setWrapText(true);
		}
		return style;
	}

	public HSSFCellStyle getOneStyle(HSSFWorkbook workbook, boolean isWarp) {
		HSSFCellStyle style = workbook.createCellStyle();
		style.setBorderLeft((short) 1); // 左边框
		style.setBorderRight((short) 1); // 右边框
		style.setBorderBottom((short) 1);
		style.setBorderTop((short) 1);
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style.setDataFormat(cellFormat);
		if (isWarp) {
			style.setWrapText(true);
		}
		return style;
	}

	private Map<String, HSSFCellStyle> createStyles(HSSFWorkbook workbook) {
		Map<String, HSSFCellStyle> map = new HashMap<String, HSSFCellStyle>();
		map.put("one", getOneStyle(workbook, false));
		map.put("oneWrap", getOneStyle(workbook, true));
		map.put("two", getTwoStyle(workbook, false));
		map.put("twoWrap", getTwoStyle(workbook, true));
		return map;
	}

	public CellStyle getStyles(Map<String, HSSFCellStyle> map, boolean needOne,
			boolean isWrap) {
		if (needOne && isWrap) {
			return map.get("oneWrap");
		}
		if (needOne) {
			return map.get("one");
		}
		if (needOne == false && isWrap) {
			return map.get("twoWrap");
		}
		return map.get("two");
	}

	public void createSheetForMap(HSSFWorkbook workbook, ExportParams entity,
			List<ExcelExportEntity> entityList,
			Collection<? extends Map<?, ?>> dataSet) {
		if (workbook == null || entity == null || entityList == null
				|| dataSet == null) {
			throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
		}
		Sheet sheet = null;
		try {
			sheet = workbook.createSheet(entity.getSheetName());
		} catch (Exception e) {
			// 重复遍历,出现了重名现象,创建非指定的名称Sheet
			sheet = workbook.createSheet();
		}
		try {
			dataHanlder = entity.getDataHanlder();
			if (dataHanlder != null) {
				needHanlderList = Arrays.asList(dataHanlder
						.getNeedHandlerFields());
			}
			// 创建表格属性
			Map<String, HSSFCellStyle> styles = createStyles(workbook);
			Drawing patriarch = sheet.createDrawingPatriarch();
			List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
			if (entity.isAddIndex()) {
				excelParams.add(indexExcelEntity());
			}
			excelParams.addAll(entityList);
			sortAllParams(excelParams);
			int index = createHeaderAndTitle(entity, sheet, workbook,
					excelParams);
			int titleHeight = index;
			setCellWith(excelParams, sheet);
			short rowHeight = getRowHeight(excelParams);
			setCurrentIndex(1);
			Iterator<?> its = dataSet.iterator();
			List<Object> tempList = new ArrayList<Object>();
			while (its.hasNext()) {
				Object t = its.next();
				index += createCells(patriarch, index, t, excelParams, sheet,
						workbook, styles, rowHeight);
				tempList.add(t);
				if (index >= MAX_NUM)
					break;
			}
			mergeCells(sheet, excelParams, titleHeight);

			its = dataSet.iterator();
			for (int i = 0, le = tempList.size(); i < le; i++) {
				its.next();
				its.remove();
			}
			// 发现还有剩余list 继续循环创建Sheet
			if (dataSet.size() > 0) {
				createSheetForMap(workbook, entity, entityList, dataSet);
			}

		} catch (Exception e) {
			e.printStackTrace();
			logger.error(e.getMessage(),e.fillInStackTrace());
			throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR,
					e.getCause());
		}
	}

}
