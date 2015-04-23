package org.jeecgframework.poi.excel.export.template;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.cache.ExcelCache;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.export.base.ExcelExportBase;
import org.jeecgframework.poi.excel.export.styler.IExcelExportStyler;
import org.jeecgframework.poi.exception.excel.ExcelExportException;
import org.jeecgframework.poi.exception.excel.enums.ExcelExportEnum;
import org.jeecgframework.poi.util.POIPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Excel 导出根据模板导出
 * 
 * @author JueYue
 * @date 2013-10-17
 * @version 1.0
 */
public final class ExcelExportOfTemplateUtil extends ExcelExportBase {

    private static final Logger LOGGER            = LoggerFactory
                                                      .getLogger(ExcelExportOfTemplateUtil.class);

    private static final String START_STR         = "{{";
    private static final String END_STR           = "}}";
    private static final String NUMBER_Symbol     = "N:";
    /**
     * 缓存temp 的for each创建的cell ,跳过这个cell的模板语法查找,提高效率
     */
    private Set<String>         tempCreateCellSet = new HashSet<String>();

    /**
     * 往Sheet 填充正常数据,根据表头信息 使用导入的部分逻辑,坐对象映射
     * 
     * @param params
     * @param pojoClass
     * @param dataSet
     * @param workbook
     */
    private void addDataToSheet(TemplateExportParams params, Class<?> pojoClass,
                                Collection<?> dataSet, Sheet sheet, Workbook workbook)
                                                                                      throws Exception {

        if (workbook instanceof XSSFWorkbook) {
            super.type = ExcelType.XSSF;
        }
        // 获取表头数据
        Map<String, Integer> titlemap = getTitleMap(params, sheet);
        Drawing patriarch = sheet.createDrawingPatriarch();
        // 得到所有字段
        Field[] fileds = POIPublicUtil.getClassFields(pojoClass);
        ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
        String targetId = null;
        if (etarget != null) {
            targetId = etarget.value();
        }
        // 创建表格样式
        setExcelExportStyler((IExcelExportStyler) params.getStyle().getConstructor(Workbook.class)
            .newInstance(workbook));
        // 获取实体对象的导出数据
        List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
        getAllExcelField(null, targetId, fileds, excelParams, pojoClass, null);
        if (excelParams.size() == 0) {
            return;
        }
        // 根据表头进行筛选排序
        sortAndFilterExportField(excelParams, titlemap);
        short rowHeight = getRowHeight(excelParams);
        int index = params.getHeadingRows() + params.getHeadingStartRow(), titleHeight = index;
        //下移数据,模拟插入
        sheet.shiftRows(params.getHeadingRows() + params.getHeadingStartRow(),
            sheet.getLastRowNum(), getShiftRows(dataSet, excelParams), true, true);

        Iterator<?> its = dataSet.iterator();
        while (its.hasNext()) {
            Object t = its.next();
            index += createCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight);
        }
        // 合并同类项
        mergeCells(sheet, excelParams, titleHeight);
    }

    /**
     * 下移数据
     * @param its
     * @param excelParams
     * @return
     */
    private int getShiftRows(Collection<?> dataSet, List<ExcelExportEntity> excelParams)
                                                                                        throws Exception {
        int size = 0;
        Iterator<?> its = dataSet.iterator();
        while (its.hasNext()) {
            Object t = its.next();
            size += getOneObjectSize(t, excelParams);
        }
        return size;
    }

    /**
     * 获取单个对象的高度,主要是处理一堆多的情况
     * 
     * @param styles
     * @param rowHeight
     * @throws Exception
     */
    public int getOneObjectSize(Object t, List<ExcelExportEntity> excelParams) throws Exception {
        ExcelExportEntity entity;
        int maxHeight = 1;
        for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
            entity = excelParams.get(k);
            if (entity.getList() != null) {
                Collection<?> list = (Collection<?>) entity.getMethod().invoke(t, new Object[] {});
                if (list != null && list.size() > maxHeight) {
                    maxHeight = list.size();
                }
            }
        }
        return maxHeight;

    }

    public Workbook createExcleByTemplate(TemplateExportParams params, Class<?> pojoClass,
                                          Collection<?> dataSet, Map<String, Object> map) {
        // step 1. 判断模板的地址
        if (params == null || map == null || StringUtils.isEmpty(params.getTemplateUrl())) {
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
            for (int i = 0; i < params.getSheetNum().length; i++) {
                parseTemplate(wb.getSheetAt(i), map);
            }
            if (dataSet != null) {
                // step 4. 正常的数据填充
                dataHanlder = params.getDataHanlder();
                if (dataHanlder != null) {
                    needHanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
                }
                addDataToSheet(params, pojoClass, dataSet, wb.getSheetAt(0), wb);
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
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
    private Workbook getCloneWorkBook(TemplateExportParams params) throws Exception {
        return ExcelCache.getWorkbook(params.getTemplateUrl(), params.getSheetNum());

    }

    /**
     * 获取参数值
     * 
     * @param params
     * @param map
     * @return
     */
    private String getParamsValue(String params, Map<String, Object> map) throws Exception {
        if (params.indexOf(".") != -1) {
            String[] paramsArr = params.split("\\.");
            return String
                .valueOf(POIPublicUtil.getValueDoWhile(map.get(paramsArr[0]), paramsArr, 1));
        }
        return map.containsKey(params) ? map.get(params).toString() : "";
    }

    /**
     * 获取表头数据,设置表头的序号
     * 
     * @param params
     * @param sheet
     * @return
     */
    private Map<String, Integer> getTitleMap(TemplateExportParams params, Sheet sheet) {
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

    private void parseTemplate(Sheet sheet, Map<String, Object> map) throws Exception {
        Row row = null;
        int index = 0;
        while (index <= sheet.getLastRowNum()) {
            row = sheet.getRow(index++);
            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                if (!tempCreateCellSet.contains(row.getRowNum() + "_"
                                                + row.getCell(i).getColumnIndex())) {
                    setValueForCellByMap(row.getCell(i), map);
                }
            }
        }
    }

    /**
     * 给每个Cell通过解析方式set值
     * 
     * @param cell
     * @param map
     */
    private void setValueForCellByMap(Cell cell, Map<String, Object> map) throws Exception {
        int cellType = cell.getCellType();
        if (cellType != Cell.CELL_TYPE_STRING && cellType == Cell.CELL_TYPE_NUMERIC) {
            return;
        }
        String oldString;
        cell.setCellType(Cell.CELL_TYPE_STRING);
        oldString = cell.getStringCellValue();
        if (oldString != null && oldString.indexOf(START_STR) != -1
            && !oldString.contains("foreach||")) {
            // setp 2. 判断是否含有解析函数
            String params;
            boolean isNumber = false;
            if (oldString.indexOf(NUMBER_Symbol) != -1) {
                isNumber = true;
                oldString = oldString.substring(2);
            }
            while (oldString.indexOf(START_STR) != -1) {
                params = oldString.substring(oldString.indexOf(START_STR) + 2,
                    oldString.indexOf(END_STR));

                oldString = oldString.replace(START_STR + params + END_STR,
                    getParamsValue(params.trim(), map));
            }
            //如何是数值 类型,就按照数值类型进行设置
            if (isNumber) {
                cell.setCellValue(Double.parseDouble(oldString));
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
            } else {
                cell.setCellValue(oldString);
            }
        }
        //判断foreach 这种方法
        if (oldString != null
            && (oldString.trim().startsWith("foreach||") || oldString.trim().startsWith(
                "!foreach||"))) {
            addListDataToExcel(cell, map, oldString.trim());
        }
    }

    /**
     * 利用foreach循环输出数据
     * @param cell 
     * @param map
     * @param oldString
     * @throws Exception 
     */
    private void addListDataToExcel(Cell cell, Map<String, Object> map, String name)
                                                                                    throws Exception {
        boolean isCreate = !name.startsWith("!");
        Collection<?> datas = (Collection<?>) map.get(name.substring(name.indexOf("||") + 2,
            name.indexOf(START_STR)));
        List<String> columns = getAllDataColumns(cell, name);
        if (datas == null) {
            return;
        }
        Iterator<?> its = datas.iterator();
        Row row;
        int rowIndex = cell.getRow().getRowNum() + 1;
        //处理当前行
        if (its.hasNext()) {
            Object t = its.next();
            setForEeachCellValue(isCreate, cell.getRow(), cell.getColumnIndex(), t, columns);
        }
        while (its.hasNext()) {
            Object t = its.next();
            if (isCreate) {
                row = cell.getRow().getSheet().createRow(rowIndex++);
            } else {
                row = cell.getRow().getSheet().getRow(rowIndex++);
                if (row == null) {
                    row = cell.getRow().getSheet().createRow(rowIndex - 1);
                }
            }
            setForEeachCellValue(isCreate, row, cell.getColumnIndex(), t, columns);
        }
    }

    private void setForEeachCellValue(boolean isCreate, Row row, int columnIndex, Object t,
                                      List<String> columns) throws Exception {
        for (int i = 0, max = columnIndex + columns.size(); i < max; i++) {
            if (row.getCell(i) == null)
                row.createCell(i);
        }
        for (int i = 0, max = columns.size(); i < max; i++) {
            String val = String.valueOf(POIPublicUtil.getValueDoWhile(t, columns.get(i)
                .split("\\."), 0));
            row.getCell(i + columnIndex).setCellValue(val);
            tempCreateCellSet.add(row.getRowNum() + "_" + (i + columnIndex));
        }

    }

    /**
     * 获取迭代的数据的值
     * @param cell
     * @param name
     * @return
     */
    private List<String> getAllDataColumns(Cell cell, String name) {
        List<String> columns = new ArrayList<String>();
        if (name.contains(END_STR)) {
            columns.add(name.substring(name.indexOf(START_STR) + 2, name.indexOf(END_STR)).trim());
            cell.setCellValue("");
            return columns;
        }
        columns.add(name.substring(name.indexOf(START_STR) + 2).trim());
        int index = cell.getColumnIndex();
        Cell tempCell;
        while (true) {
            tempCell = cell.getRow().getCell(++index);
            String cellStringString;
            try {//不允许为空
                cellStringString = tempCell.getStringCellValue();
                if (StringUtils.isBlank(cellStringString)) {
                    throw new ExcelExportException();
                }
            } catch (Exception e) {
                throw new ExcelExportException("for each 当中存在空字符串,请检查模板");
            }
            //把读取过的cell 置为空
            cell.setCellValue("");
            if (cellStringString.contains(END_STR)) {
                columns.add(cellStringString.trim().replace(END_STR, ""));
                break;
            } else {
                columns.add(cellStringString.trim());
            }

        }
        return columns;
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
            if (excelParams.get(i).getList() != null && excelParams.get(i).getList().size() > 0) {
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

}
