/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package cn.afterturn.easypoi.excel.export.template;

import cn.afterturn.easypoi.cache.ExcelCache;
import cn.afterturn.easypoi.entity.ImageEntity;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.TemplateSumEntity;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.entity.params.ExcelForEachParams;
import cn.afterturn.easypoi.excel.export.base.BaseExportService;
import cn.afterturn.easypoi.excel.export.styler.IExcelExportStyler;
import cn.afterturn.easypoi.excel.html.helper.MergedRegionHelper;
import cn.afterturn.easypoi.exception.excel.ExcelExportException;
import cn.afterturn.easypoi.exception.excel.enums.ExcelExportEnum;
import cn.afterturn.easypoi.util.*;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.*;

import static cn.afterturn.easypoi.excel.ExcelExportUtil.SHEET_NAME;
import static cn.afterturn.easypoi.util.PoiElUtil.*;

/**
 * Excel 导出根据模板导出
 *
 * @author JueYue
 * 2013-10-17
 * @version 1.0
 */
public final class ExcelExportOfTemplateUtil extends BaseExportService {

    private static final Logger LOGGER = LoggerFactory
            .getLogger(ExcelExportOfTemplateUtil.class);

    /**
     * 缓存TEMP 的for each创建的cell ,跳过这个cell的模板语法查找,提高效率
     */
    private Set<String>          tempCreateCellSet = new HashSet<String>();
    /**
     * 模板参数,全局都用到
     */
    private TemplateExportParams templateParams;
    /**
     * 单元格合并信息
     */
    private MergedRegionHelper   mergedRegionHelper;

    private TemplateSumHandler templateSumHandler;

    /**
     * 往Sheet 填充正常数据,根据表头信息 使用导入的部分逻辑,坐对象映射
     *
     * @param sheet
     * @param pojoClass
     * @param dataSet
     * @param workbook
     */
    private void addDataToSheet(Class<?> pojoClass, Collection<?> dataSet, Sheet sheet,
                                Workbook workbook) throws Exception {

        // 获取表头数据
        Map<String, Integer> titlemap  = getTitleMap(sheet);
        Drawing              patriarch = PoiExcelGraphDataUtil.getDrawingPatriarch(sheet);
        // 得到所有字段
        Field[]     fileds   = PoiPublicUtil.getClassFields(pojoClass);
        ExcelTarget etarget  = pojoClass.getAnnotation(ExcelTarget.class);
        String      targetId = null;
        if (etarget != null) {
            targetId = etarget.value();
        }
        // 获取实体对象的导出数据
        List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
        getAllExcelField(null, targetId, fileds, excelParams, pojoClass, null, null);
        // 根据表头进行筛选排序
        sortAndFilterExportField(excelParams, titlemap);
        short rowHeight = getRowHeight(excelParams);
        int index = templateParams.getHeadingRows() + templateParams.getHeadingStartRow(),
                titleHeight = index;
        int shiftRows = getShiftRows(dataSet, excelParams);
        //下移数据,模拟插入
        sheet.shiftRows(templateParams.getHeadingRows() + templateParams.getHeadingStartRow(),
                sheet.getLastRowNum(), shiftRows, true, true);
        mergedRegionHelper.shiftRows(sheet, templateParams.getHeadingRows() + templateParams.getHeadingStartRow(), shiftRows,
                sheet.getLastRowNum() - templateParams.getHeadingRows() - templateParams.getHeadingStartRow());
        templateSumHandler.shiftRows(templateParams.getHeadingRows() + templateParams.getHeadingStartRow(), shiftRows);
        PoiExcelTempUtil.reset(sheet, templateParams.getHeadingRows() + templateParams.getHeadingStartRow(), sheet.getLastRowNum());
        if (excelParams.size() == 0) {
            return;
        }
        Iterator<?> its = dataSet.iterator();
        while (its.hasNext()) {
            Object t = its.next();
            index += createCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight, 0)[0];
        }
        // 合并同类项
        mergeCells(sheet, excelParams, titleHeight);
    }

    /**
     * 利用foreach循环输出数据
     *
     * @param cell
     * @param map
     * @param name
     * @throws Exception
     */
    private void addListDataToExcel(Cell cell, Map<String, Object> map,
                                    String name) throws Exception {
        boolean isCreate = !name.contains(FOREACH_NOT_CREATE);
        boolean isShift  = name.contains(FOREACH_AND_SHIFT);
        name = name.replace(FOREACH_NOT_CREATE, EMPTY).replace(FOREACH_AND_SHIFT, EMPTY)
                .replace(FOREACH, EMPTY).replace(START_STR, EMPTY);
        String[]      keys  = name.replaceAll("\\s{1,}", " ").trim().split(" ");
        Collection<?> datas = (Collection<?>) PoiPublicUtil.getParamsValue(keys[0], map);
        Object[] columnsInfo = getAllDataColumns(cell, name.replace(keys[0], EMPTY),
                mergedRegionHelper);
        if (datas == null) {
            return;
        }
        Iterator<?> its     = datas.iterator();
        int         rowspan = (Integer) columnsInfo[0], colspan = (Integer) columnsInfo[1];
        @SuppressWarnings("unchecked")
        List<ExcelForEachParams> columns = (List<ExcelForEachParams>) columnsInfo[2];
        Row row      = null;
        int rowIndex = cell.getRow().getRowNum() + 1;
        //处理当前行
        int loopSize = 0;
        if (its.hasNext()) {
            Object t = its.next();
            loopSize = setForeachRowCellValue(isCreate, cell.getRow(), cell.getColumnIndex(), t, columns, map,
                    rowspan, colspan, mergedRegionHelper)[0];
            rowIndex += rowspan - 1 + loopSize - 1;
        }
        //修复不论后面有没有数据,都应该执行的是插入操作
        if (isShift && datas.size() * rowspan > 1 && cell.getRowIndex() + rowspan < cell.getRow().getSheet().getLastRowNum()) {
            int lastRowNum = cell.getRow().getSheet().getLastRowNum();
            int shiftRows  = lastRowNum - cell.getRowIndex() - rowspan;
            cell.getRow().getSheet().shiftRows(cell.getRowIndex() + rowspan, lastRowNum, (datas.size() - 1) * rowspan, true, true);
            mergedRegionHelper.shiftRows(cell.getSheet(), cell.getRowIndex() + rowspan, (datas.size() - 1) * rowspan, shiftRows);
            templateSumHandler.shiftRows(cell.getRowIndex() + rowspan, (datas.size() - 1) * rowspan);
            PoiExcelTempUtil.reset(cell.getSheet(), cell.getRowIndex() + rowspan + (datas.size() - 1) * rowspan, cell.getRow().getSheet().getLastRowNum());
        }
        while (its.hasNext()) {
            Object t = its.next();
            row = createRow(rowIndex, cell.getSheet(), isCreate, rowspan);
            loopSize = setForeachRowCellValue(isCreate, row, cell.getColumnIndex(), t, columns, map, rowspan,
                    colspan, mergedRegionHelper)[0];
            rowIndex += rowspan + loopSize - 1;
        }
    }

    /**
     * 下移数据
     *
     * @param dataSet
     * @param excelParams
     * @return
     */
    private int getShiftRows(Collection<?> dataSet,
                             List<ExcelExportEntity> excelParams) throws Exception {
        int         size = 0;
        Iterator<?> its  = dataSet.iterator();
        while (its.hasNext()) {
            Object t = its.next();
            size += getOneObjectSize(t, excelParams);
        }
        return size;
    }

    /**
     * 获取单个对象的高度,主要是处理一堆多的情况
     *
     * @param t
     * @param excelParams
     * @throws Exception
     */
    private int getOneObjectSize(Object t, List<ExcelExportEntity> excelParams) throws Exception {
        ExcelExportEntity entity;
        int               maxHeight = 1;
        for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
            entity = excelParams.get(k);
            if (entity.getList() != null) {
                Collection<?> list = (Collection<?>) entity.getMethod().invoke(t, new Object[]{});
                if (list != null && list.size() > maxHeight) {
                    maxHeight = list.size();
                }
            }
        }
        return maxHeight;

    }

    public Workbook createExcelCloneByTemplate(TemplateExportParams params,
                                               Map<Integer, List<Map<String, Object>>> map) {
        // step 1. 判断模板的地址
        if (params == null || map == null || StringUtils.isEmpty(params.getTemplateUrl())) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        Workbook wb = null;
        // step 2. 判断模板的Excel类型,解析模板
        try {
            this.templateParams = params;
            wb = ExcelCache.getWorkbook(templateParams.getTemplateUrl(), templateParams.getSheetNum(),
                    true);
            int          oldSheetNum  = wb.getNumberOfSheets();
            List<String> oldSheetName = new ArrayList<>();
            for (int i = 0; i < oldSheetNum; i++) {
                oldSheetName.add(wb.getSheetName(i));
            }
            // 把所有的KEY排个顺序
            List<Map<String, Object>> mapList;
            List<Integer>             sheetNumList = new ArrayList<>();
            sheetNumList.addAll(map.keySet());
            Collections.sort(sheetNumList);

            //把需要克隆的全部克隆一遍
            for (Integer sheetNum : sheetNumList) {
                mapList = map.get(sheetNum);
                for (int i = mapList.size(); i > 0; i--) {
                    wb.cloneSheet(sheetNum);
                }
            }
            for (int i = 0; i < oldSheetName.size(); i++) {
                wb.removeSheetAt(wb.getSheetIndex(oldSheetName.get(i)));
            }
            // 创建表格样式
            setExcelExportStyler((IExcelExportStyler) templateParams.getStyle()
                    .getConstructor(Workbook.class).newInstance(wb));
            // step 3. 解析模板
            int sheetIndex = 0;
            for (Integer sheetNum : sheetNumList) {
                mapList = map.get(sheetNum);
                for (int i = mapList.size() - 1; i >= 0; i--) {
                    tempCreateCellSet.clear();
                    if (mapList.get(i).containsKey(SHEET_NAME)) {
                        wb.setSheetName(sheetIndex, mapList.get(i).get(SHEET_NAME).toString());
                    }
                    parseTemplate(wb.getSheetAt(sheetIndex), mapList.get(i), params.isColForEach());
                    if (params.isReadonly()) {
                        wb.getSheetAt(i).protectSheet(UUID.randomUUID().toString());
                    }
                    sheetIndex++;
                }
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            return null;
        }
        return wb;
    }

    public Workbook createExcelByTemplate(TemplateExportParams params,
                                          Map<Integer, Map<String, Object>> map) {
        // step 1. 判断模板的地址
        if (params == null || map == null || StringUtils.isEmpty(params.getTemplateUrl())) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        Workbook wb = null;
        // step 2. 判断模板的Excel类型,解析模板
        try {
            this.templateParams = params;
            wb = getCloneWorkBook();
            // 创建表格样式
            setExcelExportStyler((IExcelExportStyler) templateParams.getStyle()
                    .getConstructor(Workbook.class).newInstance(wb));
            // step 3. 解析模板
            for (int i = 0, le = params.isScanAllsheet() ? wb.getNumberOfSheets()
                    : params.getSheetNum().length; i < le; i++) {
                if (params.getSheetName() != null && params.getSheetName().length > i
                        && StringUtils.isNotEmpty(params.getSheetName()[i])) {
                    wb.setSheetName(i, params.getSheetName()[i]);
                }
                tempCreateCellSet.clear();
                parseTemplate(wb.getSheetAt(i), map.get(i), params.isColForEach());
                if (params.isReadonly()) {
                    wb.getSheetAt(i).protectSheet(UUID.randomUUID().toString());
                }
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            return null;
        }
        return wb;
    }

    public Workbook createExcelByTemplate(TemplateExportParams params, Class<?> pojoClass,
                                          Collection<?> dataSet, Map<String, Object> map) {
        // step 1. 判断模板的地址
        if (params == null || map == null || (StringUtils.isEmpty(params.getTemplateUrl()) && params.getTemplateWb() == null)) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        Workbook wb = null;
        // step 2. 判断模板的Excel类型,解析模板
        try {
            this.templateParams = params;
            if (params.getTemplateWb() != null) {
                wb = params.getTemplateWb();
            } else {
                wb = getCloneWorkBook();
            }
            if (params.getDictHandler() != null) {
                this.dictHandler = params.getDictHandler();
            }
            if (params.getI18nHandler() != null) {
                this.i18nHandler = params.getI18nHandler();
            }
            // 创建表格样式
            setExcelExportStyler((IExcelExportStyler) templateParams.getStyle()
                    .getConstructor(Workbook.class).newInstance(wb));
            // step 3. 解析模板
            for (int i = 0, le = params.isScanAllsheet() ? wb.getNumberOfSheets()
                    : params.getSheetNum().length; i < le; i++) {
                if (params.getSheetName() != null && params.getSheetName().length > i
                        && StringUtils.isNotEmpty(params.getSheetName()[i])) {
                    wb.setSheetName(i, params.getSheetName()[i]);
                }
                tempCreateCellSet.clear();
                parseTemplate(wb.getSheetAt(i), map, params.isColForEach());
                if (params.isReadonly()) {
                    wb.getSheetAt(i).protectSheet(UUID.randomUUID().toString());
                }
            }
            if (dataSet != null) {
                // step 4. 正常的数据填充
                dataHandler = params.getDataHandler();
                if (dataHandler != null) {
                    needHandlerList = Arrays.asList(dataHandler.getNeedHandlerFields());
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
     * @throws Exception
     */
    private Workbook getCloneWorkBook() throws Exception {
        return ExcelCache.getWorkbook(templateParams.getTemplateUrl(), templateParams.getSheetNum(),
                templateParams.isScanAllsheet());

    }

    /**
     * 获取表头数据,设置表头的序号
     *
     * @param sheet
     * @return
     */
    private Map<String, Integer> getTitleMap(Sheet sheet) {
        Row                  row      = null;
        Iterator<Cell>       cellTitle;
        Map<String, Integer> titlemap = new HashMap<String, Integer>();
        for (int j = 0; j < templateParams.getHeadingRows(); j++) {
            row = sheet.getRow(j + templateParams.getHeadingStartRow());
            cellTitle = row.cellIterator();
            int i = row.getFirstCellNum();
            while (cellTitle.hasNext()) {
                Cell   cell  = cellTitle.next();
                String value = cell.getStringCellValue();
                if (!StringUtils.isEmpty(value)) {
                    titlemap.put(value, i);
                }
                i = i + 1;
            }
        }
        return titlemap;

    }

    private void parseTemplate(Sheet sheet, Map<String, Object> map,
                               boolean colForeach) throws Exception {
        if (sheet.getWorkbook() instanceof XSSFWorkbook) {
            super.type = ExcelType.XSSF;
        }
        deleteCell(sheet, map);
        mergedRegionHelper = new MergedRegionHelper(sheet);
        templateSumHandler = new TemplateSumHandler(sheet);
        if (colForeach) {
            colForeach(sheet, map);
        }
        Row row   = null;
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

        //修改需要处理的统计值
        handlerSumCell(sheet);
    }

    private void handlerSumCell(Sheet sheet) {
        for (TemplateSumEntity sumEntity : templateSumHandler.getDataList()) {
            Cell cell = sheet.getRow(sumEntity.getRow()).getCell(sumEntity.getCol());
            if (cell.getStringCellValue().contains(sumEntity.getSumKey())) {
                cell.setCellValue(cell.getStringCellValue()
                        .replace("sum:(" + sumEntity.getSumKey() + ")", sumEntity.getValue() + ""));
            } else {
                cell.setCellValue(cell.getStringCellValue() + sumEntity.getValue());
            }
        }
    }

    /**
     * 先进行列的循环,因为涉及很多数据
     *
     * @param sheet
     * @param map
     */
    private void colForeach(Sheet sheet, Map<String, Object> map) throws Exception {
        Row  row   = null;
        Cell cell  = null;
        int  index = 0;
        while (index <= sheet.getLastRowNum()) {
            row = sheet.getRow(index++);
            if (row == null) {
                continue;
            }
            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                cell = row.getCell(i);
                if (row.getCell(i) != null && (cell.getCellType() == CellType.STRING
                        || cell.getCellType() == CellType.NUMERIC)) {
                    String text = PoiCellUtil.getCellValue(cell);
                    if (text.contains(FOREACH_COL) || text.contains(FOREACH_COL_VALUE)) {
                        foreachCol(cell, map, text);
                    }
                }
            }
        }
    }

    /**
     * 循环列表
     *
     * @param cell
     * @param map
     * @param name
     * @throws Exception
     */
    private void foreachCol(Cell cell, Map<String, Object> map, String name) throws Exception {
        boolean isCreate = name.contains(FOREACH_COL_VALUE);
        name = name.replace(FOREACH_COL_VALUE, EMPTY).replace(FOREACH_COL, EMPTY).replace(START_STR,
                EMPTY);
        String[]      keys  = name.replaceAll("\\s{1,}", " ").trim().split(" ");
        Collection<?> datas = (Collection<?>) PoiPublicUtil.getParamsValue(keys[0], map);
        Object[] columnsInfo = getAllDataColumns(cell, name.replace(keys[0], EMPTY),
                mergedRegionHelper);
        if (datas == null) {
            return;
        }
        Iterator<?> its     = datas.iterator();
        int         rowspan = (Integer) columnsInfo[0], colspan = (Integer) columnsInfo[1];
        @SuppressWarnings("unchecked")
        List<ExcelForEachParams> columns = (List<ExcelForEachParams>) columnsInfo[2];
        while (its.hasNext()) {
            Object t = its.next();
            setForeachRowCellValue(true, cell.getRow(), cell.getColumnIndex(), t, columns, map,
                    rowspan, colspan, mergedRegionHelper);
            if (cell.getRow().getCell(cell.getColumnIndex() + colspan) == null) {
                cell.getRow().createCell(cell.getColumnIndex() + colspan);
            }
            cell = cell.getRow().getCell(cell.getColumnIndex() + colspan);
        }
        if (isCreate) {
            cell = cell.getRow().getCell(cell.getColumnIndex() - 1);
            cell.setCellValue(cell.getStringCellValue() + END_STR);
        }
    }

    /**
     * 先判断删除,省得影响效率
     *
     * @param sheet
     * @param map
     * @throws Exception
     */
    private void deleteCell(Sheet sheet, Map<String, Object> map) throws Exception {
        Row  row   = null;
        Cell cell  = null;
        int  index = 0;
        while (index <= sheet.getLastRowNum()) {
            row = sheet.getRow(index++);
            if (row == null) {
                continue;
            }
            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                cell = row.getCell(i);
                if (row.getCell(i) != null && (cell.getCellType() == CellType.STRING
                        || cell.getCellType() == CellType.NUMERIC)) {
                    cell.setCellType(CellType.STRING);
                    String text = cell.getStringCellValue();
                    if (text.contains(IF_DELETE)) {
                        if (Boolean.valueOf(
                                eval(text.substring(text.indexOf(START_STR) + 2, text.indexOf(END_STR))
                                        .trim(), map).toString())) {
                            PoiSheetUtil.deleteColumn(sheet, i);
                            i--;
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
        CellType cellType = cell.getCellType();
        if (cellType != CellType.STRING && cellType != CellType.NUMERIC) {
            return;
        }
        String oldString;
        oldString = cell.getStringCellValue();
        if (oldString != null && oldString.indexOf(START_STR) != -1
                && !oldString.contains(FOREACH)) {
            // step 2. 判断是否含有解析函数
            boolean isNumber = false;
            if (isHasSymbol(oldString, NUMBER_SYMBOL)) {
                isNumber = true;
                oldString = oldString.replaceFirst(NUMBER_SYMBOL, "");
            }
            boolean isStyleBySelf = false;
            if (isHasSymbol(oldString, STYLE_SELF)) {
                isStyleBySelf = true;
                oldString = oldString.replaceFirst(STYLE_SELF, "");
            }
            boolean isDict = false;
            String  dict   = null;
            if (isHasSymbol(oldString, DICT_HANDLER)) {
                isDict = true;
                dict = oldString.substring(oldString.indexOf(DICT_HANDLER) + 5).split(";")[0];
                oldString = oldString.replaceFirst(DICT_HANDLER, "");
            }
            boolean isI18n = false;
            if (isHasSymbol(oldString, I18N_HANDLER)) {
                isI18n = true;
                oldString = oldString.replaceFirst(I18N_HANDLER, "");
            }
            Object obj = PoiPublicUtil.getRealValue(oldString, map);
            if (isDict) {
                obj = dictHandler.toName(dict, null, oldString, obj);
            }
            if (isI18n) {
                obj = i18nHandler.getLocaleName(obj.toString());
            }
            //如何是数值 类型,就按照数值类型进行设置// 如果是图片就设置为图片
            if (obj instanceof ImageEntity) {
                ImageEntity img = (ImageEntity) obj;
                cell.setCellValue("");
                if (img.getRowspan() > 1 || img.getColspan() > 1) {
                    img.setHeight(0);
                    PoiMergeCellUtil.addMergedRegion(cell.getSheet(), cell.getRowIndex(),
                            cell.getRowIndex() + img.getRowspan() - 1, cell.getColumnIndex(), cell.getColumnIndex() + img.getColspan() - 1);
                }
                createImageCell(cell, img.getHeight(), img.getRowspan(), img.getColspan(), img.getUrl(), img.getData());
            } else if (isNumber && StringUtils.isNotBlank(obj.toString())) {
                cell.setCellValue(Double.parseDouble(obj.toString()));
            } else {
                cell.setCellValue(obj.toString());
            }
        }
        //判断foreach 这种方法
        if (oldString != null && oldString.contains(FOREACH)) {
            addListDataToExcel(cell, map, oldString.trim());
        }

    }

    private boolean isHasSymbol(String text, String symbol) {
        return text.startsWith(symbol) || text.contains("{" + symbol)
                || text.contains(" " + symbol);
    }

    /**
     * 创建并返回第一个Row
     *
     * @param sheet
     * @param rowIndex
     * @param isCreate
     * @param rows
     * @return
     */
    private Row createRow(int rowIndex, Sheet sheet, boolean isCreate, int rows) {
        for (int i = 0; i < rows; i++) {
            if (isCreate) {
                sheet.createRow(rowIndex++);
            } else if (sheet.getRow(rowIndex++) == null) {
                sheet.createRow(rowIndex - 1);
            }
        }
        return sheet.getRow(rowIndex - rows);
    }

    /**
     * 循环迭代创建,遍历row
     *
     * @param isCreate
     * @param row
     * @param columnIndex
     * @param t
     * @param columns
     * @param map
     * @param rowspan
     * @param colspan
     * @param mergedRegionHelper
     * @return rowSize, cellSize
     * @throws Exception
     */
    private int[] setForeachRowCellValue(boolean isCreate, Row row, int columnIndex, Object t,
                                         List<ExcelForEachParams> columns, Map<String, Object> map,
                                         int rowspan, int colspan,
                                         MergedRegionHelper mergedRegionHelper) throws Exception {
        createRowCellSetStyle(row, columnIndex, columns, rowspan, colspan);
        //填写数据
        ExcelForEachParams params;
        int                loopSize = 1;
        int                loopCi   = 1;
        row = row.getSheet().getRow(row.getRowNum() - rowspan + 1);
        for (int k = 0; k < rowspan; k++) {
            int ci = columnIndex;
            row.setHeight(getMaxHeight(k, colspan, columns));
            for (int i = 0; i < colspan && i < columns.size(); i++) {
                boolean isNumber = false;
                params = columns.get(colspan * k + i);
                tempCreateCellSet.add(row.getRowNum() + "_" + (ci));
                if (params == null) {
                    continue;
                }
                if (StringUtils.isEmpty(params.getName())
                        && StringUtils.isEmpty(params.getConstValue())) {
                    row.getCell(ci).setCellStyle(params.getCellStyle());
                    ci = ci + params.getColspan();
                    continue;
                }
                String val;
                Object obj = null;
                //是不是常量
                String tempStr = params.getName();
                if (StringUtils.isEmpty(params.getName())) {
                    val = params.getConstValue();
                } else {
                    if (isHasSymbol(tempStr, NUMBER_SYMBOL)) {
                        isNumber = true;
                        tempStr = tempStr.replaceFirst(NUMBER_SYMBOL, "");
                    }
                    map.put(templateParams.getTempParams(), t);
                    boolean isDict = false;
                    String  dict   = null;
                    if (isHasSymbol(tempStr, DICT_HANDLER)) {
                        isDict = true;
                        dict = tempStr.substring(tempStr.indexOf(DICT_HANDLER) + 5).split(";")[0];
                        tempStr = tempStr.replaceFirst(DICT_HANDLER, "");
                        tempStr = tempStr.replaceFirst(dict + ";", "");
                    }
                    obj = eval(tempStr, map);
                    if (isDict && !(obj instanceof Collection)) {
                        obj = dictHandler.toName(dict, t, tempStr, obj);
                    }
                    val = obj.toString();
                }
                if (obj != null && obj instanceof Collection) {
                    // 需要找到哪一级别是集合 ,方便后面的replace
                    String collectName = evalFindName(tempStr, map);
                    int[] loop = setForEachLoopRowCellValue(row, ci, (Collection) obj, columns,
                            params, map, rowspan, colspan, mergedRegionHelper, collectName);
                    loopSize = Math.max(loopSize, loop[0]);
                    i += loop[1] - 1;
                    ci = loop[2] - params.getColspan();
                } else if (obj != null && obj instanceof ImageEntity) {
                    ImageEntity img = (ImageEntity) obj;
                    row.getCell(ci).setCellValue("");
                    if (img.getRowspan() > 1 || img.getColspan() > 1) {
                        img.setHeight(0);
                        row.getCell(ci).getSheet().addMergedRegion(new CellRangeAddress(row.getCell(ci).getRowIndex(),
                                row.getCell(ci).getRowIndex() + img.getRowspan() - 1, row.getCell(ci).getColumnIndex(), row.getCell(ci).getColumnIndex() + img.getColspan() - 1));
                    }
                    createImageCell(row.getCell(ci), img.getHeight(), img.getRowspan(), img.getColspan(), img.getUrl(), img.getData());
                } else if (isNumber && StringUtils.isNotEmpty(val)) {
                    row.getCell(ci).setCellValue(Double.parseDouble(val));
                } else {
                    try {
                        row.getCell(ci).setCellValue(val);
                    } catch (Exception e) {
                        LOGGER.error(e.getMessage(), e);
                    }
                }
                if (params.getCellStyle() != null) {
                    row.getCell(ci).setCellStyle(params.getCellStyle());
                }
                //判断这个属性是不是需要统计
                if (params.isNeedSum()) {
                    templateSumHandler.addValueOfKey(params.getName(), val);
                }
                //如果合并单元格,就把这个单元格的样式和之前的保持一致
                setMergedRegionStyle(row, ci, params);
                //合并对应单元格
                if ((params.getRowspan() != 1 || params.getColspan() != 1)
                        && !mergedRegionHelper.isMergedRegion(row.getRowNum() + 1, ci)
                        && PoiCellUtil.isMergedRegion(row.getSheet(), row.getRowNum(), ci)) {
                    PoiMergeCellUtil.addMergedRegion(row.getSheet(), row.getRowNum(),
                            row.getRowNum() + params.getRowspan() - 1, ci,
                            ci + params.getColspan() - 1);
                }
                ci = ci + params.getColspan();
            }
            loopCi = Math.max(loopCi, ci);
            // 需要把需要合并的单元格合并了  --- 不是集合的栏位合并了
            if (loopSize > 1) {
                handlerLoopMergedRegion(row, columnIndex, columns, loopSize);
            }
            row = row.getSheet().getRow(row.getRowNum() + 1);
        }
        return new int[]{loopSize, loopCi};
    }

    /**
     * 迭代把不是集合的数据都合并了
     *
     * @param row
     * @param columnIndex
     * @param columns
     * @param loopSize
     */
    private void handlerLoopMergedRegion(Row row, int columnIndex, List<ExcelForEachParams> columns, int loopSize) {
        for (int i = 0; i < columns.size(); i++) {
            if (!columns.get(i).isCollectCell()) {
                PoiMergeCellUtil.addMergedRegion(row.getSheet(), row.getRowNum(),
                        row.getRowNum() + loopSize - 1, columnIndex,
                        columnIndex + columns.get(i).getColspan() - 1);
            }
            columnIndex = columnIndex + columns.get(i).getColspan();
        }
    }

    private short getMaxHeight(int k, int colspan, List<ExcelForEachParams> columns) {
        short high = columns.get(0).getHeight();
        int   n    = k;
        while (n > 0) {
            if (columns.get(n * colspan).getHeight() == 0) {
                n--;
            } else {
                high = columns.get(n * colspan).getHeight();
                break;
            }
        }
        return high;
    }

    /**
     * 处理内循环
     *
     * @param row
     * @param columnIndex
     * @param obj
     * @param columns
     * @param params
     * @param map
     * @param rowspan
     * @param colspan
     * @param mergedRegionHelper
     * @param collectName
     * @return [rowNums, columnsNums, ciIndex]
     * @throws Exception
     */
    private int[] setForEachLoopRowCellValue(Row row, int columnIndex, Collection obj, List<ExcelForEachParams> columns,
                                             ExcelForEachParams params, Map<String, Object> map,
                                             int rowspan, int colspan,
                                             MergedRegionHelper mergedRegionHelper, String collectName) throws Exception {

        //多个一起遍历 -去掉第一层 把所有的数据遍历一遍
        //STEP 1拿到所有的和当前一样项目的字段
        List<ExcelForEachParams> temp    = getLoopEachParams(columns, columnIndex, collectName);
        Iterator<?>              its     = obj.iterator();
        Row                      tempRow = row;
        int                      nums    = 0;
        int                      ci      = columnIndex;
        while (its.hasNext()) {
            Object data = its.next();
            map.put("loop_" + columnIndex, data);
            int[] loopArr = setForeachRowCellValue(false, tempRow, columnIndex, data, temp, map, rowspan,
                    colspan, mergedRegionHelper);
            nums += loopArr[0];
            ci = Math.max(ci, loopArr[1]);
            map.remove("loop_" + columnIndex);
            tempRow = createRow(tempRow.getRowNum() + loopArr[0], row.getSheet(), false, rowspan);
        }
        for (int i = 0; i < temp.size(); i++) {
            temp.get(i).setName(temp.get(i).getTempName().pop());
            //都是集合
            temp.get(i).setCollectCell(true);

        }
        return new int[]{nums, temp.size(), ci};
    }

    /**
     * 根据 当前是集合的信息,把后面整个集合的迭代获取出来,并替换掉集合的前缀方便后面取数
     *
     * @param columns
     * @param columnIndex
     * @param collectName
     * @return
     */
    private List<ExcelForEachParams> getLoopEachParams(List<ExcelForEachParams> columns, int columnIndex, String collectName) {
        List<ExcelForEachParams> temp = new ArrayList<>();
        for (int i = 0; i < columns.size(); i++) {
            //先置为不是集合
            columns.get(i).setCollectCell(false);
            if (columns.get(i) == null || columns.get(i).getName().contains(collectName)) {
                temp.add(columns.get(i));
                if (columns.get(i).getTempName() == null) {
                    columns.get(i).setTempName(new Stack<>());
                }
                columns.get(i).setCollectCell(true);
                columns.get(i).getTempName().push(columns.get(i).getName());
                columns.get(i).setName(columns.get(i).getName().replace(collectName, "loop_" + columnIndex));
            }
        }
        return temp;
    }

    private void createRowCellSetStyle(Row row, int columnIndex, List<ExcelForEachParams> columns,
                                       int rowspan, int colspan) {
        //所有的cell创建一遍
        for (int i = 0; i < rowspan; i++) {
            int size = columns.size();
            for (int j = columnIndex, max = columnIndex + colspan; j < max; j++) {
                if (row.getCell(j) == null) {
                    row.createCell(j);
                    CellStyle style = row.getRowNum() % 2 == 0
                            ? getStyles(false,
                            size <= j - columnIndex ? null : columns.get(j - columnIndex))
                            : getStyles(true,
                            size <= j - columnIndex ? null : columns.get(j - columnIndex));
                    //返回的styler不为空时才使用,否则使用Excel设置的,更加推荐Excel设置的样式
                    if (style != null) {
                        row.getCell(j).setCellStyle(style);
                    }
                }

            }
            if (i < rowspan - 1) {
                row = row.getSheet().getRow(row.getRowNum() + 1);
            }
        }
    }

    private CellStyle getStyles(boolean isSingle, ExcelForEachParams excelForEachParams) {
        return excelExportStyler.getTemplateStyles(isSingle, excelForEachParams);
    }

    /**
     * 设置合并单元格的样式
     *
     * @param row
     * @param ci
     * @param params
     */
    private void setMergedRegionStyle(Row row, int ci, ExcelForEachParams params) {
        //第一行数据
        for (int i = 1; i < params.getColspan(); i++) {
            if (params.getCellStyle() != null) {
                row.getCell(ci + i).setCellStyle(params.getCellStyle());
            }
        }
        for (int i = 1; i < params.getRowspan(); i++) {
            for (int j = 0; j < params.getColspan(); j++) {
                if (params.getCellStyle() != null) {
                    row.getCell(ci + j).setCellStyle(params.getCellStyle());
                }
            }
        }
    }

    /**
     * 获取迭代的数据的值
     *
     * @param cell
     * @param name
     * @param mergedRegionHelper
     * @return
     */
    private Object[] getAllDataColumns(Cell cell, String name,
                                       MergedRegionHelper mergedRegionHelper) {
        List<ExcelForEachParams> columns = new ArrayList<ExcelForEachParams>();
        cell.setCellValue("");
        columns.add(getExcelTemplateParams(name.replace(END_STR, EMPTY), cell, mergedRegionHelper));
        int rowspan = 1, colspan = 1;
        if (!name.contains(END_STR)) {
            int index = cell.getColumnIndex();
            //保存col 的开始列
            int startIndex = cell.getColumnIndex();
            Row row        = cell.getRow();
            while (index < row.getLastCellNum()) {
                int colSpan = columns.get(columns.size() - 1) != null
                        ? columns.get(columns.size() - 1).getColspan() : 1;
                index += colSpan;


                for (int i = 1; i < colSpan; i++) {
                    //添加合并的单元格,这些单元可能不是空,但是没有值,所以也需要跳过
                    columns.add(null);
                    continue;
                }
                cell = row.getCell(index);
                //可能是合并的单元格
                if (cell == null) {
                    //读取是判断,跳过
                    columns.add(null);
                    continue;
                }
                String cellStringString;
                try {//不允许为空 便利单元格必须有结尾和值
                    cellStringString = cell.getStringCellValue();
                    if (StringUtils.isBlank(cellStringString) && colspan + startIndex <= index) {
                        throw new ExcelExportException("for each 当中存在空字符串,请检查模板");
                    } else if (StringUtils.isBlank(cellStringString)
                            && colspan + startIndex > index) {
                        //读取是判断,跳过,数据为空,但是不是第一次读这一列,所以可以跳过
                        columns.add(new ExcelForEachParams(null, cell.getCellStyle(), (short) 0));
                        continue;
                    }
                } catch (Exception e) {
                    throw new ExcelExportException(ExcelExportEnum.TEMPLATE_ERROR, e);
                }
                //把读取过的cell 置为空
                cell.setCellValue("");
                if (cellStringString.contains(END_STR)) {
                    columns.add(getExcelTemplateParams(cellStringString.replace(END_STR, EMPTY),
                            cell, mergedRegionHelper));
                    //补全缺失的cell(合并单元格后面的)
                    int lastCellColspan = columns.get(columns.size() - 1).getColspan();
                    for (int i = 1; i < lastCellColspan; i++) {
                        //添加合并的单元格,这些单元可能不是空,但是没有值,所以也需要跳过
                        columns.add(null);
                    }
                    break;
                } else if (cellStringString.contains(WRAP)) {
                    columns.add(getExcelTemplateParams(cellStringString.replace(WRAP, EMPTY), cell,
                            mergedRegionHelper));
                    //发现换行符,执行换行操作
                    colspan = index - startIndex + 1;
                    index = startIndex - columns.get(columns.size() - 1).getColspan();
                    row = row.getSheet().getRow(row.getRowNum() + 1);
                    rowspan++;
                } else {
                    columns.add(getExcelTemplateParams(cellStringString.replace(WRAP, EMPTY), cell,
                            mergedRegionHelper));
                }
            }
        }
        colspan = 0;
        for (int i = 0; i < columns.size(); i++) {
            colspan += columns.get(i) != null ? columns.get(i).getColspan() : 0;
        }
        colspan = colspan / rowspan;
        return new Object[]{rowspan, colspan, columns};
    }

    /**
     * 获取模板参数
     *
     * @param name
     * @param cell
     * @param mergedRegionHelper
     * @return
     */
    private ExcelForEachParams getExcelTemplateParams(String name, Cell cell,
                                                      MergedRegionHelper mergedRegionHelper) {
        name = name.trim();
        ExcelForEachParams params = new ExcelForEachParams(name, cell.getCellStyle(),
                cell.getRow().getHeight());
        //判断是不是常量
        if (name.startsWith(CONST) && name.endsWith(CONST)) {
            params.setName(null);
            params.setConstValue(name.substring(1, name.length() - 1));
        }
        //判断是不是空
        if (NULL.equals(name)) {
            params.setName(null);
            params.setConstValue(EMPTY);
        }
        //获取合并单元格的数据
        if (mergedRegionHelper.isMergedRegion(cell.getRowIndex() + 1, cell.getColumnIndex())) {
            Integer[] colAndrow = mergedRegionHelper.getRowAndColSpan(cell.getRowIndex() + 1,
                    cell.getColumnIndex());
            params.setRowspan(colAndrow[0]);
            params.setColspan(colAndrow[1]);
        }
        params.setNeedSum(templateSumHandler.isSumKey(params.getName()));
        return params;
    }

    /**
     * 对导出序列进行排序和塞选
     *
     * @param excelParams
     * @param titlemap
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