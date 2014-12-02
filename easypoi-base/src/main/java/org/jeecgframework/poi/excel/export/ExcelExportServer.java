package org.jeecgframework.poi.excel.export;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

    private final static Logger logger     = LoggerFactory.getLogger(ExcelExportServer.class);

    private static final short  cellFormat = (short) BuiltinFormats.getBuiltinFormat("TEXT");

    // 最大行数,超过自动多Sheet
    private int                 MAX_NUM    = 60000;

    private int createHeaderAndTitle(ExportParams entity, Sheet sheet, Workbook workbook,
                                     List<ExcelExportEntity> excelParams) {
        int rows = 0, feildWidth = getFieldWidth(excelParams);
        if (entity.getTitle() != null) {
            rows += createHeaderRow(entity, sheet, workbook, feildWidth);
        }
        rows += createTitleRow(entity, sheet, workbook, rows, excelParams);
        sheet.createFreezePane(0, rows, 0, rows);
        return rows;
    }

    /**
     * 创建 表头改变
     * 
     * @param entity
     * @param sheet
     * @param workbook
     * @param feildWidth
     */
    public int createHeaderRow(ExportParams entity, Sheet sheet, Workbook workbook, int feildWidth) {
        Row row = sheet.createRow(0);
        row.setHeight(entity.getTitleHeight());
        createStringCell(row, 0, entity.getTitle(), getHeaderStyle(workbook, entity), null);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, feildWidth));
        if (entity.getSecondTitle() != null) {
            row = sheet.createRow(1);
            row.setHeight(entity.getSecondTitleHeight());
            CellStyle style = workbook.createCellStyle();
            style.setAlignment(CellStyle.ALIGN_RIGHT);
            createStringCell(row, 0, entity.getSecondTitle(), style, null);
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, feildWidth));
            return 2;
        }
        return 1;
    }

    public void createSheet(Workbook workbook, ExportParams entity, Class<?> pojoClass,
                            Collection<?> dataSet, String type) {
        if (logger.isDebugEnabled()) {
            logger.debug("Excel export start ,class is {}", pojoClass);
            logger.debug("Excel version is {}", type.equals(PoiBaseConstants.HSSF) ? "03" : "07");
        }
        if (workbook == null || entity == null || pojoClass == null || dataSet == null) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        if (type.equals(PoiBaseConstants.XSSF)) {
            MAX_NUM = 1000000;
        }
        super.type = type;
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
                needHanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
            }
            // 创建表格属性
            Map<String, CellStyle> styles = createStyles(workbook);
            Drawing patriarch = sheet.createDrawingPatriarch();
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity());
            }
            // 得到所有字段
            Field fileds[] = POIPublicUtil.getClassFields(pojoClass);
            ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
            String targetId = etarget == null ? null : etarget.value();
            getAllExcelField(entity.getExclusions(), targetId, fileds, excelParams, pojoClass, null);
            sortAllParams(excelParams);
            int index = createHeaderAndTitle(entity, sheet, workbook, excelParams);
            int titleHeight = index;
            setCellWith(excelParams, sheet);
            short rowHeight = getRowHeight(excelParams);
            setCurrentIndex(1);
            Iterator<?> its = dataSet.iterator();
            List<Object> tempList = new ArrayList<Object>();
            while (its.hasNext()) {
                Object t = its.next();
                index += createCells(patriarch, index, t, excelParams, sheet, workbook, styles,
                    rowHeight);
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
                createSheet(workbook, entity, pojoClass, dataSet, type);
            }

        } catch (Exception e) {
            logger.error(e.getMessage(), e.fillInStackTrace());
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
        }
    }

    public void createSheetForMap(Workbook workbook, ExportParams entity,
                                  List<ExcelExportEntity> entityList,
                                  Collection<? extends Map<?, ?>> dataSet, String type) {
        if (logger.isDebugEnabled()) {
            logger.debug("Excel version is {}", type.equals(PoiBaseConstants.HSSF) ? "03" : "07");
        }
        if (workbook == null || entity == null || entityList == null || dataSet == null) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        if (type.equals(PoiBaseConstants.XSSF)) {
            MAX_NUM = 1000000;
        }
        super.type = type;
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
                needHanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
            }
            // 创建表格属性
            Map<String, CellStyle> styles = createStyles(workbook);
            Drawing patriarch = sheet.createDrawingPatriarch();
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity());
            }
            excelParams.addAll(entityList);
            sortAllParams(excelParams);
            int index = createHeaderAndTitle(entity, sheet, workbook, excelParams);
            int titleHeight = index;
            setCellWith(excelParams, sheet);
            short rowHeight = getRowHeight(excelParams);
            setCurrentIndex(1);
            Iterator<?> its = dataSet.iterator();
            List<Object> tempList = new ArrayList<Object>();
            while (its.hasNext()) {
                Object t = its.next();
                index += createCells(patriarch, index, t, excelParams, sheet, workbook, styles,
                    rowHeight);
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
                createSheetForMap(workbook, entity, entityList, dataSet, type);
            }

        } catch (Exception e) {
            e.printStackTrace();
            logger.error(e.getMessage(), e.fillInStackTrace());
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
        }
    }

    private Map<String, CellStyle> createStyles(Workbook workbook) {
        Map<String, CellStyle> map = new HashMap<String, CellStyle>();
        map.put("one", getOneStyle(workbook, false));
        map.put("oneWrap", getOneStyle(workbook, true));
        map.put("two", getTwoStyle(workbook, false));
        map.put("twoWrap", getTwoStyle(workbook, true));
        return map;
    }

    /**
     * 创建表头
     * 
     * @param title
     * @param index
     */
    private int createTitleRow(ExportParams title, Sheet sheet, Workbook workbook, int index,
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
            createStringCell(row, cellIndex, entity.getName(), titleStyle, entity);
            if (entity.getList() != null) {
                List<ExcelExportEntity> sTitel = entity.getList();
                sheet.addMergedRegion(new CellRangeAddress(index, index, cellIndex, cellIndex
                                                                                    + sTitel.size()
                                                                                    - 1));
                for (int j = 0, size = sTitel.size(); j < size; j++) {
                    createStringCell(listRow, cellIndex, sTitel.get(j).getName(), titleStyle,
                        entity);
                    cellIndex++;
                }
            } else if (rows == 2) {
                sheet.addMergedRegion(new CellRangeAddress(index, index + 1, cellIndex, cellIndex));
            }
            cellIndex++;
        }
        return rows;

    }

    /**
     * 表明的Style
     * 
     * @param workbook
     * @return
     */
    public CellStyle getHeaderStyle(Workbook workbook, ExportParams entity) {
        CellStyle titleStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 24);
        titleStyle.setFont(font);
        titleStyle.setFillForegroundColor(entity.getColor());
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        return titleStyle;
    }

    public CellStyle getOneStyle(Workbook workbook, boolean isWarp) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderLeft((short) 1); // 左边框
        style.setBorderRight((short) 1); // 右边框
        style.setBorderBottom((short) 1);
        style.setBorderTop((short) 1);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setDataFormat(cellFormat);
        if (isWarp) {
            style.setWrapText(true);
        }
        return style;
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

    public CellStyle getStyles(Map<String, CellStyle> map, boolean needOne, boolean isWrap) {
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

    /**
     * 字段说明的Style
     * 
     * @param workbook
     * @return
     */
    public CellStyle getTitleStyle(Workbook workbook, ExportParams entity) {
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFillForegroundColor(entity.getHeaderColor()); // 填充的背景颜色
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND); // 填充图案
        titleStyle.setWrapText(true);
        return titleStyle;
    }

    public CellStyle getTwoStyle(Workbook workbook, boolean isWarp) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderLeft((short) 1); // 左边框
        style.setBorderRight((short) 1); // 右边框
        style.setBorderBottom((short) 1);
        style.setBorderTop((short) 1);
        style.setFillForegroundColor((short) 41); // 填充的背景颜色
        style.setFillPattern(CellStyle.SOLID_FOREGROUND); // 填充图案
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setDataFormat(cellFormat);
        if (isWarp) {
            style.setWrapText(true);
        }
        return style;
    }

    private ExcelExportEntity indexExcelEntity() {
        ExcelExportEntity entity = new ExcelExportEntity();
        entity.setOrderNum(0);
        entity.setName("序号");
        entity.setWidth(10);
        entity.setFormat(PoiBaseConstants.IS_ADD_INDEX);
        return entity;
    }

}
