/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 *   
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
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
import org.jeecgframework.poi.excel.entity.params.ExcelTemplateParams;
import org.jeecgframework.poi.excel.export.base.ExcelExportBase;
import org.jeecgframework.poi.excel.export.styler.IExcelExportStyler;
import org.jeecgframework.poi.exception.excel.ExcelExportException;
import org.jeecgframework.poi.exception.excel.enums.ExcelExportEnum;

import static org.jeecgframework.poi.util.PoiElUtil.*;

import org.jeecgframework.poi.util.PoiPublicUtil;
import org.jeecgframework.poi.util.PoiSheetUtility;
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

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelExportOfTemplateUtil.class);

    /**
     * 缓存TEMP 的for each创建的cell ,跳过这个cell的模板语法查找,提高效率
     */
    private Set<String>          tempCreateCellSet = new HashSet<String>();
    /**
     * 模板参数,全局都用到
     */
    private TemplateExportParams teplateParams;

    /**
     * 往Sheet 填充正常数据,根据表头信息 使用导入的部分逻辑,坐对象映射
     * 
     * @param teplateParams
     * @param pojoClass
     * @param dataSet
     * @param workbook
     */
    private void addDataToSheet(Class<?> pojoClass, Collection<?> dataSet, Sheet sheet,
                                Workbook workbook) throws Exception {

        if (workbook instanceof XSSFWorkbook) {
            super.type = ExcelType.XSSF;
        }
        // 获取表头数据
        Map<String, Integer> titlemap = getTitleMap(sheet);
        Drawing patriarch = sheet.createDrawingPatriarch();
        // 得到所有字段
        Field[] fileds = PoiPublicUtil.getClassFields(pojoClass);
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
        int index = teplateParams.getHeadingRows() + teplateParams.getHeadingStartRow(),
                titleHeight = index;
        //下移数据,模拟插入
        sheet.shiftRows(teplateParams.getHeadingRows() + teplateParams.getHeadingStartRow(),
            sheet.getLastRowNum(), getShiftRows(dataSet, excelParams), true, true);
        if (excelParams.size() == 0) {
            return;
        }
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
    private int getShiftRows(Collection<?> dataSet,
                             List<ExcelExportEntity> excelParams) throws Exception {
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
            this.teplateParams = params;
            wb = getCloneWorkBook();
            // 创建表格样式
            setExcelExportStyler((IExcelExportStyler) teplateParams.getStyle()
                .getConstructor(Workbook.class).newInstance(wb));
            // step 3. 解析模板
            for (int i = 0, le = params.isScanAllsheet() ? wb.getNumberOfSheets()
                : params.getSheetNum().length; i < le; i++) {
                if (params.getSheetName() != null && params.getSheetName().length > i
                    && StringUtils.isNotEmpty(params.getSheetName()[i])) {
                    wb.setSheetName(i, params.getSheetName()[i]);
                }
                tempCreateCellSet.clear();
                parseTemplate(wb.getSheetAt(i), map);
            }
            if (dataSet != null) {
                // step 4. 正常的数据填充
                dataHanlder = params.getDataHanlder();
                if (dataHanlder != null) {
                    needHanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
                }
                addDataToSheet(pojoClass, dataSet, wb.getSheetAt(params.getDataSheetNum()), wb);
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
     * @param teplateParams
     * @throws Exception
     * @Author JueYue
     * @date 2013-11-11
     */
    private Workbook getCloneWorkBook() throws Exception {
        return ExcelCache.getWorkbook(teplateParams.getTemplateUrl(), teplateParams.getSheetNum(),
            teplateParams.isScanAllsheet());

    }

    /**
     * 获取表头数据,设置表头的序号
     * 
     * @param teplateParams
     * @param sheet
     * @return
     */
    private Map<String, Integer> getTitleMap(Sheet sheet) {
        Row row = null;
        Iterator<Cell> cellTitle;
        Map<String, Integer> titlemap = new HashMap<String, Integer>();
        for (int j = 0; j < teplateParams.getHeadingRows(); j++) {
            row = sheet.getRow(j + teplateParams.getHeadingStartRow());
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
        deleteCell(sheet, map);
        Row row = null;
        int index = 0;
        while (index <= sheet.getLastRowNum()) {
            row = sheet.getRow(index++);
            if (row == null) {
                continue;
            }
            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                if (row.getCell(i) != null && !tempCreateCellSet
                    .contains(row.getRowNum() + "_" + row.getCell(i).getColumnIndex())) {
                    setValueForCellByMap(row.getCell(i), map);
                }
            }
        }
    }

    /**
     * 先判断删除,省得影响效率
     * @param sheet
     * @param map
     * @throws Exception 
     */
    private void deleteCell(Sheet sheet, Map<String, Object> map) throws Exception {
        Row row = null;
        Cell cell = null;
        int index = 0;
        while (index <= sheet.getLastRowNum()) {
            row = sheet.getRow(index++);
            if (row == null) {
                continue;
            }
            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                cell = row.getCell(i);
                if (row.getCell(i) != null && (cell.getCellType() == Cell.CELL_TYPE_STRING
                                               || cell.getCellType() == Cell.CELL_TYPE_NUMERIC)) {
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    String text = cell.getStringCellValue();
                    if (text.contains(IF_DELETE)) {
                        if (Boolean.valueOf(
                            eval(text.substring(text.indexOf(START_STR) + 2, text.indexOf(END_STR))
                                .trim(), map).toString())) {
                            PoiSheetUtility.deleteColumn(sheet, i);
                        }
                        cell.setCellValue("");
                    }
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
        if (cellType != Cell.CELL_TYPE_STRING && cellType != Cell.CELL_TYPE_NUMERIC) {
            return;
        }
        String oldString;
        cell.setCellType(Cell.CELL_TYPE_STRING);
        oldString = cell.getStringCellValue();
        if (oldString != null && oldString.indexOf(START_STR) != -1
            && !oldString.contains(FOREACH)) {
            // step 2. 判断是否含有解析函数
            String params = null;
            boolean isNumber = false;
            if (isNumber(oldString)) {
                isNumber = true;
                oldString = oldString.replace(NUMBER_SYMBOL, "");
            }
            while (oldString.indexOf(START_STR) != -1) {
                params = oldString.substring(oldString.indexOf(START_STR) + 2,
                    oldString.indexOf(END_STR));

                oldString = oldString.replace(START_STR + params + END_STR,
                    eval(params, map).toString());
            }
            //如何是数值 类型,就按照数值类型进行设置
            if (isNumber && StringUtils.isNotBlank(oldString)) {
                cell.setCellValue(Double.parseDouble(oldString));
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
            } else {
                cell.setCellValue(oldString);
            }
        }
        //判断foreach 这种方法
        if (oldString != null && oldString.contains(FOREACH)) {
            addListDataToExcel(cell, map, oldString.trim());
        }

    }

    private boolean isNumber(String text) {
        return text.startsWith(NUMBER_SYMBOL) || text.contains("{" + NUMBER_SYMBOL)
               || text.contains(" " + NUMBER_SYMBOL);
    }

    /**
     * 利用foreach循环输出数据
     * @param cell 
     * @param map
     * @param oldString
     * @throws Exception 
     */
    private void addListDataToExcel(Cell cell, Map<String, Object> map,
                                    String name) throws Exception {
        boolean isCreate = !name.contains(FOREACH_NOT_CREATE);
        boolean isShift = name.contains(FOREACH_AND_SHIFT);
        name = name.replace(FOREACH_NOT_CREATE, EMPTY).replace(FOREACH_AND_SHIFT, EMPTY)
            .replace(FOREACH, EMPTY).replace(START_STR, EMPTY);
        String[] keys = name.replaceAll("\\s{1,}", " ").trim().split(" ");
        Collection<?> datas = (Collection<?>) PoiPublicUtil.getParamsValue(keys[0], map);
        List<ExcelTemplateParams> columns = getAllDataColumns(cell, name.replace(keys[0], EMPTY));
        if (datas == null) {
            return;
        }
        Iterator<?> its = datas.iterator();
        Row row;
        int rowIndex = cell.getRow().getRowNum() + 1;
        //处理当前行
        if (its.hasNext()) {
            Object t = its.next();
            cell.getRow().setHeight(columns.get(0).getHeight());
            setForEeachCellValue(isCreate, cell.getRow(), cell.getColumnIndex(), t, columns, map);
        }
        if (isShift) {
            cell.getRow().getSheet().shiftRows(cell.getRowIndex() + 1,
                cell.getRow().getSheet().getLastRowNum(), datas.size() - 1, true, true);
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
            row.setHeight(columns.get(0).getHeight());
            setForEeachCellValue(isCreate, row, cell.getColumnIndex(), t, columns, map);
        }
    }

    private void setForEeachCellValue(boolean isCreate, Row row, int columnIndex, Object t,
                                      List<ExcelTemplateParams> columns,
                                      Map<String, Object> map) throws Exception {
        for (int i = 0, max = columnIndex + columns.size(); i < max; i++) {
            if (row.getCell(i) == null)
                row.createCell(i);
        }
        for (int i = 0, max = columns.size(); i < max; i++) {
            boolean isNumber = false;
            String tempStr = new String(columns.get(i).getName());
            if (isNumber(tempStr)) {
                isNumber = true;
                tempStr = tempStr.replace(NUMBER_SYMBOL, "");
            }
            map.put(teplateParams.getTempParams(), t);
            String val = eval(tempStr, map).toString();
            if (isNumber && StringUtils.isNotEmpty(val)) {
                row.getCell(i + columnIndex).setCellValue(Double.parseDouble(val));
                row.getCell(i + columnIndex).setCellType(Cell.CELL_TYPE_NUMERIC);
            } else {
                row.getCell(i + columnIndex).setCellValue(val);
            }
            row.getCell(i + columnIndex).setCellStyle(columns.get(i).getCellStyle());
            tempCreateCellSet.add(row.getRowNum() + "_" + (i + columnIndex));
        }

    }

    /**
     * 获取迭代的数据的值
     * @param cell
     * @param name
     * @return
     */
    private List<ExcelTemplateParams> getAllDataColumns(Cell cell, String name) {
        List<ExcelTemplateParams> columns = new ArrayList<ExcelTemplateParams>();
        cell.setCellValue("");
        if (name.contains(END_STR)) {
            columns.add(new ExcelTemplateParams(name.replace(END_STR, EMPTY).trim(),
                cell.getCellStyle(), cell.getRow().getHeight()));
            return columns;
        }
        columns.add(
            new ExcelTemplateParams(name.trim(), cell.getCellStyle(), cell.getRow().getHeight()));
        int index = cell.getColumnIndex();
        Cell tempCell;
        while (true) {
            tempCell = cell.getRow().getCell(++index);
            if (tempCell == null) {
                break;
            }
            String cellStringString;
            try {//允许为空,单表示已经完结了,因为可能被删除了
                cellStringString = tempCell.getStringCellValue();
                if (StringUtils.isBlank(cellStringString)) {
                    break;
                }
            } catch (Exception e) {
                throw new ExcelExportException("for each 当中存在空字符串,请检查模板");
            }
            //把读取过的cell 置为空
            tempCell.setCellValue("");
            if (cellStringString.contains(END_STR)) {
                columns.add(new ExcelTemplateParams(cellStringString.trim().replace(END_STR, ""),
                    tempCell.getCellStyle(), tempCell.getRow().getHeight()));
                break;
            } else {
                if (cellStringString.trim().contains(teplateParams.getTempParams())) {
                    columns.add(new ExcelTemplateParams(cellStringString.trim(),
                        tempCell.getCellStyle(), tempCell.getRow().getHeight()));
                } else {
                    //最后一行被删除了
                    break;
                }
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

    public Workbook createExcleByTemplate(TemplateExportParams params,
                                          Map<Integer, Map<String, Object>> map) {
        // step 1. 判断模板的地址
        if (params == null || map == null || StringUtils.isEmpty(params.getTemplateUrl())) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        Workbook wb = null;
        // step 2. 判断模板的Excel类型,解析模板
        try {
            this.teplateParams = params;
            wb = getCloneWorkBook();
            // step 3. 解析模板
            for (int i = 0, le = params.isScanAllsheet() ? wb.getNumberOfSheets()
                : params.getSheetNum().length; i < le; i++) {
                if (params.getSheetName() != null && params.getSheetName().length > i
                    && StringUtils.isNotEmpty(params.getSheetName()[i])) {
                    wb.setSheetName(i, params.getSheetName()[i]);
                }
                tempCreateCellSet.clear();
                parseTemplate(wb.getSheetAt(i), map.get(i));
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            return null;
        }
        return wb;
    }

}
