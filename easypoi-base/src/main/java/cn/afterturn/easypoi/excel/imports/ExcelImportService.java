/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License. You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software distributed under the License
 * is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
 * or implied. See the License for the specific language governing permissions and limitations under
 * the License.
 */
package cn.afterturn.easypoi.excel.imports;

import cn.afterturn.easypoi.handler.inter.IExcelDataModel;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.lang.reflect.Field;
import java.util.*;

import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.params.ExcelCollectionParams;
import cn.afterturn.easypoi.excel.entity.params.ExcelImportEntity;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import cn.afterturn.easypoi.excel.entity.result.ExcelVerifyHandlerResult;
import cn.afterturn.easypoi.excel.imports.base.ImportBaseService;
import cn.afterturn.easypoi.exception.excel.ExcelImportException;
import cn.afterturn.easypoi.exception.excel.enums.ExcelImportEnum;
import cn.afterturn.easypoi.handler.inter.IExcelModel;
import cn.afterturn.easypoi.util.PoiCellUtil;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import cn.afterturn.easypoi.util.PoiReflectorUtil;
import cn.afterturn.easypoi.util.PoiValidationUtil;

/**
 * Excel 导入服务
 *
 * @author JueYue 2014年6月26日 下午9:20:51
 */
@SuppressWarnings({"rawtypes", "unchecked", "hiding"})
public class ExcelImportService extends ImportBaseService {

    private final static Logger LOGGER = LoggerFactory.getLogger(ExcelImportService.class);

    private CellValueService cellValueServer;

    private boolean verfiyFail = false;
    /**
     * 异常数据styler
     */
    private CellStyle errorCellStyle;

    private List<Row> successRow;
    private List<Row> failRow;
    private List failCollection = new ArrayList();

    public ExcelImportService() {
        successRow = new ArrayList<Row>();
        failRow = new ArrayList<Row>();
        this.cellValueServer = new CellValueService();
    }

    /***
     * 向List里面继续添加元素
     *
     * @param object
     * @param param
     * @param row
     * @param titlemap
     * @param targetId
     * @param pictures
     * @param params
     */
    private void addListContinue(Object object, ExcelCollectionParams param, Row row,
                                 Map<Integer, String> titlemap, String targetId,
                                 Map<String, PictureData> pictures,
                                 ImportParams params, StringBuilder errorMsg) throws Exception {
        Collection collection = (Collection) PoiReflectorUtil.fromCache(object.getClass())
                .getValue(object, param.getName());
        Object entity = PoiPublicUtil.createObject(param.getType(), targetId);
        if (entity instanceof IExcelDataModel) {
            ((IExcelDataModel) entity).setRowNum(row.getRowNum());
        }
        String picId;
        // 是否需要加上这个对象
        boolean isUsed = false;
        for (int i = row.getFirstCellNum(); i < titlemap.size(); i++) {
            Cell cell = row.getCell(i);
            String titleString = (String) titlemap.get(i);
            if (param.getExcelParams().containsKey(titleString)) {
                if (param.getExcelParams().get(titleString).getType() == 2) {
                    picId = row.getRowNum() + "_" + i;
                    saveImage(object, picId, param.getExcelParams(), titleString, pictures, params);
                } else {
                    try {
                        saveFieldValue(params, entity, cell, param.getExcelParams(), titleString, row);
                    } catch (ExcelImportException e) {
                        // 如果需要去校验就忽略,这个错误,继续执行
                        if(params.isNeedVerfiy() && ExcelImportEnum.GET_VALUE_ERROR.equals(e.getType())){
                            errorMsg.append(" ").append(titleString).append(ExcelImportEnum.GET_VALUE_ERROR.getMsg());
                        }
                    }
                }
                isUsed = true;
            }
        }
        if (isUsed) {
            collection.add(entity);
        }
    }

    /**
     * 获取key的值,针对不同类型获取不同的值
     *
     * @author JueYue 2013-11-21
     */
    private String getKeyValue(Cell cell) {
        Object obj = PoiCellUtil.getCellValue(cell);
        return obj == null ? null : obj.toString().trim();
    }

    /**
     * 获取保存的真实路径
     */
    private String getSaveUrl(ExcelImportEntity excelImportEntity, Object object) throws Exception {
        String url = "";
        if (ExcelImportEntity.IMG_SAVE_PATH.equals(excelImportEntity.getSaveUrl())) {
            if (excelImportEntity.getMethods() != null
                    && excelImportEntity.getMethods().size() > 0) {
                object = getFieldBySomeMethod(excelImportEntity.getMethods(), object);
            }
            url = object.getClass().getName()
                    .split("\\.")[object.getClass().getName().split("\\.").length - 1];
            return excelImportEntity.getSaveUrl() + File.separator + url;
        }
        return excelImportEntity.getSaveUrl();
    }

    private <T> List<T> importExcel(Collection<T> result, Sheet sheet, Class<?> pojoClass,
                                    ImportParams params,
                                    Map<String, PictureData> pictures) throws Exception {
        List collection = new ArrayList();
        Map<String, ExcelImportEntity> excelParams = new HashMap<String, ExcelImportEntity>();
        List<ExcelCollectionParams> excelCollection = new ArrayList<ExcelCollectionParams>();
        String targetId = null;
        i18nHandler = params.getI18nHandler();
        if (!Map.class.equals(pojoClass)) {
            Field[] fileds = PoiPublicUtil.getClassFields(pojoClass);
            ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
            if (etarget != null) {
                targetId = etarget.value();
            }
            getAllExcelField(targetId, fileds, excelParams, excelCollection, pojoClass, null, null);
        }
        Iterator<Row> rows = sheet.rowIterator();
        for (int j = 0; j < params.getTitleRows(); j++) {
            rows.next();
        }
        Map<Integer, String> titlemap = getTitleMap(rows, params, excelCollection, excelParams);
        checkIsValidTemplate(titlemap, excelParams, params, excelCollection);
        Row row = null;
        Object object = null;
        String picId;
        int readRow = 0;
        //跳过无效行
        for (int i = 0; i < params.getStartRows(); i++) {
            rows.next();
        }
        //判断index 和集合,集合情况默认为第一列
        if (excelCollection.size() > 0 && params.getKeyIndex() == null) {
            params.setKeyIndex(0);
        }
        StringBuilder errorMsg;
        while (rows.hasNext()
                && (row == null
                || sheet.getLastRowNum() - row.getRowNum() > params.getLastOfInvalidRow())) {
            if (params.getReadRows() > 0 && readRow > params.getReadRows()) {
                break;
            }
            row = rows.next();
            // Fix 如果row为无效行时候跳出
            if (sheet.getLastRowNum() - row.getRowNum() < params.getLastOfInvalidRow()) {
                break;
            }
            errorMsg = new StringBuilder();
            // 判断是集合元素还是不是集合元素,如果是就继续加入这个集合,不是就创建新的对象
            // keyIndex 如果为空就不处理,仍然处理这一行
            if (params.getKeyIndex() != null
                    && (row.getCell(params.getKeyIndex()) == null
                    || StringUtils.isEmpty(getKeyValue(row.getCell(params.getKeyIndex()))))
                    && object != null) {
                for (ExcelCollectionParams param : excelCollection) {
                    addListContinue(object, param, row, titlemap, targetId, pictures, params, errorMsg);
                }
            } else {
                object = PoiPublicUtil.createObject(pojoClass, targetId);
                try {
                    Set<Integer> keys = titlemap.keySet();
                    for (Integer cn : keys) {
                        Cell cell = row.getCell(cn);
                        String titleString = (String) titlemap.get(cn);
                        if (excelParams.containsKey(titleString) || Map.class.equals(pojoClass)) {
                            if (excelParams.get(titleString) != null
                                    && excelParams.get(titleString).getType() == 2) {
                                picId = row.getRowNum() + "_" + cn;
                                saveImage(object, picId, excelParams, titleString, pictures,
                                        params);
                            } else {
                                try {
                                    saveFieldValue(params, object, cell, excelParams, titleString, row);
                                } catch (ExcelImportException e) {
                                    // 如果需要去校验就忽略,这个错误,继续执行
                                    if(params.isNeedVerfiy() && ExcelImportEnum.GET_VALUE_ERROR.equals(e.getType())){
                                        errorMsg.append(" ").append(titleString).append(ExcelImportEnum.GET_VALUE_ERROR.getMsg());
                                    }
                                }
                            }
                        }
                    }
                    //for (int i = row.getFirstCellNum(), le = titlemap.size(); i < le; i++) {

                    //}
                    if (object instanceof IExcelDataModel) {
                        ((IExcelDataModel) object).setRowNum(row.getRowNum());
                    }
                    for (ExcelCollectionParams param : excelCollection) {
                        addListContinue(object, param, row, titlemap, targetId, pictures, params, errorMsg);
                    }
                    if (verifyingDataValidity(object, row, params, pojoClass, errorMsg)) {
                        collection.add(object);
                    } else {
                        failCollection.add(object);
                    }
                } catch (ExcelImportException e) {
                    LOGGER.error("excel import error , row num:{},obj:{}", readRow, ReflectionToStringBuilder.toString(object));
                    if (!e.getType().equals(ExcelImportEnum.VERIFY_ERROR)) {
                        throw new ExcelImportException(e.getType(), e);
                    }
                } catch (Exception e) {
                    LOGGER.error("excel import error , row num:{},obj:{}", readRow, ReflectionToStringBuilder.toString(object));
                    throw new RuntimeException(e);
                }
            }
            readRow++;
        }
        return collection;
    }

    /**
     * 校验数据合法性
     */
    private boolean verifyingDataValidity(Object object, Row row, ImportParams params,
                                          Class<?> pojoClass, StringBuilder fieldErrorMsg) {
        boolean isAdd = true;
        Cell cell = null;
        if (params.isNeedVerfiy()) {
            String errorMsg = PoiValidationUtil.validation(object, params.getVerfiyGroup());
            if (StringUtils.isNotEmpty(errorMsg)) {
                cell = row.createCell(row.getLastCellNum());
                cell.setCellValue(errorMsg);
                if (object instanceof IExcelModel) {
                    IExcelModel model = (IExcelModel) object;
                    model.setErrorMsg(errorMsg);
                }
                isAdd = false;
                verfiyFail = true;
            }
        }
        if (params.getVerifyHandler() != null) {
            ExcelVerifyHandlerResult result = params.getVerifyHandler().verifyHandler(object);
            if (!result.isSuccess()) {
                if (cell == null) {
                    cell = row.createCell(row.getLastCellNum());
                }
                cell.setCellValue((StringUtils.isNoneBlank(cell.getStringCellValue())
                        ? cell.getStringCellValue() + "," : "") + result.getMsg());
                if (object instanceof IExcelModel) {
                    IExcelModel model = (IExcelModel) object;
                    model.setErrorMsg((StringUtils.isNoneBlank(model.getErrorMsg())
                            ? model.getErrorMsg() + "," : "") + result.getMsg());
                }
                isAdd = false;
                verfiyFail = true;
            }
        }
        if((params.isNeedVerfiy() || params.getVerifyHandler() != null) && fieldErrorMsg.length() > 0){
            if (object instanceof IExcelModel) {
                IExcelModel model = (IExcelModel) object;
                model.setErrorMsg((StringUtils.isNoneBlank(model.getErrorMsg())
                        ? model.getErrorMsg() + "," : "") + fieldErrorMsg.toString());
            }
            if (cell == null) {
                cell = row.createCell(row.getLastCellNum());
            }
            cell.setCellValue((StringUtils.isNoneBlank(cell.getStringCellValue())
                    ? cell.getStringCellValue() + "," : "")+ fieldErrorMsg.toString());
            isAdd = false;
            verfiyFail = true;
        }
        if (cell != null) {
            cell.setCellStyle(errorCellStyle);
            failRow.add(row);
        } else {
            successRow.add(row);
        }
        return isAdd;
    }

    /**
     * 获取表格字段列名对应信息
     */
    private Map<Integer, String> getTitleMap(Iterator<Row> rows, ImportParams params,
                                             List<ExcelCollectionParams> excelCollection,
                                             Map<String, ExcelImportEntity> excelParams) {
        Map<Integer, String> titlemap = new LinkedHashMap<Integer, String>();
        Iterator<Cell> cellTitle;
        String collectionName = null;
        ExcelCollectionParams collectionParams = null;
        Row row = null;
        for (int j = 0; j < params.getHeadRows(); j++) {
            row = rows.next();
            if (row == null) {
                continue;
            }
            cellTitle = row.cellIterator();
            while (cellTitle.hasNext()) {
                Cell cell = cellTitle.next();
                String value = getKeyValue(cell);
                value = value.replace("\n", "");
                int i = cell.getColumnIndex();
                //用以支持重名导入
                if (StringUtils.isNotEmpty(value)) {
                    if (titlemap.containsKey(i)) {
                        collectionName = titlemap.get(i);
                        collectionParams = getCollectionParams(excelCollection, collectionName);
                        titlemap.put(i, collectionName + "_" + value);
                    } else if (StringUtils.isNotEmpty(collectionName) && collectionParams != null
                            && collectionParams.getExcelParams()
                            .containsKey(collectionName + "_" + value)) {
                        titlemap.put(i, collectionName + "_" + value);
                    } else {
                        collectionName = null;
                        collectionParams = null;
                    }
                    if (StringUtils.isEmpty(collectionName)) {
                        titlemap.put(i, value);
                    }
                }
            }
        }

        // 处理指定列的情况
        Set<String> keys = excelParams.keySet();
        for (String key : keys) {
            if (key.startsWith("FIXED_")) {
                String[] arr = key.split("_");
                titlemap.put(Integer.parseInt(arr[1]), key);
            }
        }
        return titlemap;
    }

    /**
     * 获取这个名称对应的集合信息
     */
    private ExcelCollectionParams getCollectionParams(List<ExcelCollectionParams> excelCollection,
                                                      String collectionName) {
        for (ExcelCollectionParams excelCollectionParams : excelCollection) {
            if (collectionName.equals(excelCollectionParams.getExcelName())) {
                return excelCollectionParams;
            }
        }
        return null;
    }

    /**
     * Excel 导入 field 字段类型 Integer,Long,Double,Date,String,Boolean
     */
    public ExcelImportResult importExcelByIs(InputStream inputstream, Class<?> pojoClass,
                                             ImportParams params, boolean needMore) throws Exception {
        if (LOGGER.isDebugEnabled()) {
            LOGGER.debug("Excel import start ,class is {}", pojoClass);
        }
        List<T> result = new ArrayList<T>();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ExcelImportResult importResult;
        try {
            byte[] buffer = new byte[1024];
            int len;
            while ((len = inputstream.read(buffer)) > -1) {
                baos.write(buffer, 0, len);
            }
            baos.flush();

            InputStream userIs = new ByteArrayInputStream(baos.toByteArray());
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug("Excel clone success");
            }
            Workbook book = WorkbookFactory.create(userIs);

            boolean isXSSFWorkbook = !(book instanceof HSSFWorkbook);
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug("Workbook create success");
            }
            importResult = new ExcelImportResult();
            createErrorCellStyle(book);
            Map<String, PictureData> pictures;
            for (int i = params.getStartSheetIndex(); i < params.getStartSheetIndex()
                    + params.getSheetNum(); i++) {
                if (LOGGER.isDebugEnabled()) {
                    LOGGER.debug(" start to read excel by is ,startTime is {}", System.currentTimeMillis());
                }
                if (isXSSFWorkbook) {
                    pictures = PoiPublicUtil.getSheetPictrues07((XSSFSheet) book.getSheetAt(i),
                            (XSSFWorkbook) book);
                } else {
                    pictures = PoiPublicUtil.getSheetPictrues03((HSSFSheet) book.getSheetAt(i),
                            (HSSFWorkbook) book);
                }
                if (LOGGER.isDebugEnabled()) {
                    LOGGER.debug(" end to read excel by is ,endTime is {}", System.currentTimeMillis());
                }
                result.addAll(importExcel(result, book.getSheetAt(i), pojoClass, params, pictures));
                if (LOGGER.isDebugEnabled()) {
                    LOGGER.debug(" end to read excel list by sheet ,endTime is {}", System.currentTimeMillis());
                }
                if (params.isReadSingleCell()) {
                    readSingleCell(importResult, book.getSheetAt(i), params);
                    if (LOGGER.isDebugEnabled()) {
                        LOGGER.debug(" read Key-Value ,endTime is {}", System.currentTimeMillis());
                    }
                }
            }
            if (params.isNeedSave()) {
                saveThisExcel(params, pojoClass, isXSSFWorkbook, book);
            }
            importResult.setList(result);
            if (needMore) {
                InputStream successIs = new ByteArrayInputStream(baos.toByteArray());
                try {
                    Workbook successBook = WorkbookFactory.create(successIs);
                    importResult.setWorkbook(removeSuperfluousRows(successBook, failRow, params));
                    importResult.setFailWorkbook(removeSuperfluousRows(book, successRow, params));
                    importResult.setFailList(failCollection);
                    importResult.setVerfiyFail(verfiyFail);
                } finally {
                    successIs.close();
                }
            }
        } finally {
            IOUtils.closeQuietly(baos);
        }

        return importResult;
    }

    private Workbook removeSuperfluousRows(Workbook book, List<Row> rowList, ImportParams params) {
        for (int i = params.getStartSheetIndex(); i < params.getStartSheetIndex()
                + params.getSheetNum(); i++) {
            for (int j = rowList.size() - 1; j >= 0; j--) {
                if (rowList.get(j).getRowNum() < rowList.get(j).getSheet().getLastRowNum()) {
                    book.getSheetAt(i).shiftRows(rowList.get(j).getRowNum() + 1, rowList.get(j).getSheet().getLastRowNum(), -1);
                } else if (rowList.get(j).getRowNum() == rowList.get(j).getSheet().getLastRowNum()) {
                    book.getSheetAt(i).createRow(rowList.get(j).getRowNum() + 1);
                    book.getSheetAt(i).shiftRows(rowList.get(j).getRowNum() + 1, rowList.get(j).getSheet().getLastRowNum() + 1, -1);
                }
            }
        }
        return book;
    }

    /**
     * 按照键值对的方式取得Excel里面的数据
     */
    private void readSingleCell(ExcelImportResult result, Sheet sheet, ImportParams params) {
        if (result.getMap() == null) {
            result.setMap(new HashMap<String, Object>());
        }
        for (int i = 0; i < params.getTitleRows() + params.getHeadRows() + params.getStartRows(); i++) {
            getSingleCellValueForRow(result, sheet.getRow(i), params);
        }

        for (int i = sheet.getLastRowNum() - params.getLastOfInvalidRow(); i < sheet.getLastRowNum(); i++) {
            getSingleCellValueForRow(result, sheet.getRow(i), params);

        }
    }

    private void getSingleCellValueForRow(ExcelImportResult result, Row row, ImportParams params) {
        for (int j = row.getFirstCellNum(), le = row.getLastCellNum(); j < le; j++) {
            String text = PoiCellUtil.getCellValue(row.getCell(j));
            if (StringUtils.isNoneBlank(text) && text.endsWith(params.getKeyMark())) {
                if (result.getMap().containsKey(text)) {
                    if (result.getMap().get(text) instanceof String) {
                        List<String> list = new ArrayList<String>();
                        list.add((String) result.getMap().get(text));
                        result.getMap().put(text, list);
                    }
                    ((List) result.getMap().get(text)).add(PoiCellUtil.getCellValue(row.getCell(++j)));
                } else {
                    result.getMap().put(text, PoiCellUtil.getCellValue(row.getCell(++j)));
                }

            }

        }
    }

    /**
     * 检查是不是合法的模板
     */
    private void checkIsValidTemplate(Map<Integer, String> titlemap,
                                      Map<String, ExcelImportEntity> excelParams,
                                      ImportParams params,
                                      List<ExcelCollectionParams> excelCollection) {

        if (params.getImportFields() != null) {
            if (params.isNeedCheckOrder()) { // 同时校验列顺序

                if (params.getImportFields().length != titlemap.size()) {
                    LOGGER.error("excel列顺序不一致");
                    throw new ExcelImportException(ExcelImportEnum.IS_NOT_A_VALID_TEMPLATE);
                }
                int i = 0;
                for (String title : titlemap.values()) {
                    if (!StringUtils.equals(title, params.getImportFields()[i++])) {
                        LOGGER.error("excel列顺序不一致");
                        throw new ExcelImportException(ExcelImportEnum.IS_NOT_A_VALID_TEMPLATE);
                    }
                }
            } else {
                for (int i = 0, le = params.getImportFields().length; i < le; i++) {
                    if (!titlemap.containsValue(params.getImportFields()[i])) {
                        throw new ExcelImportException(ExcelImportEnum.IS_NOT_A_VALID_TEMPLATE);
                    }
                }
            }
        } else {
            Collection<ExcelImportEntity> collection = excelParams.values();
            for (ExcelImportEntity excelImportEntity : collection) {
                if (excelImportEntity.isImportField()
                        && !titlemap.containsValue(excelImportEntity.getName())) {
                    LOGGER.error(excelImportEntity.getName() + "必须有,但是没找到");
                    throw new ExcelImportException(ExcelImportEnum.IS_NOT_A_VALID_TEMPLATE);
                }
            }

            for (int i = 0, le = excelCollection.size(); i < le; i++) {
                ExcelCollectionParams collectionparams = excelCollection.get(i);
                collection = collectionparams.getExcelParams().values();
                for (ExcelImportEntity excelImportEntity : collection) {
                    if (excelImportEntity.isImportField() && !titlemap.containsValue(
                            collectionparams.getExcelName() + "_" + excelImportEntity.getName())) {
                        throw new ExcelImportException(ExcelImportEnum.IS_NOT_A_VALID_TEMPLATE);
                    }
                }
            }
        }
    }

    /**
     * 保存字段值(获取值,校验值,追加错误信息)
     */
    private void saveFieldValue(ImportParams params, Object object, Cell cell,
                                Map<String, ExcelImportEntity> excelParams, String titleString,
                                Row row) throws Exception {
        Object value = cellValueServer.getValue(params.getDataHandler(), object, cell, excelParams,
                titleString, params.getDictHandler());
        if (object instanceof Map) {
            if (params.getDataHandler() != null) {
                params.getDataHandler().setMapValue((Map) object, titleString, value);
            } else {
                ((Map) object).put(titleString, value);
            }
        } else {
            setValues(excelParams.get(titleString), object, value);
        }
    }

    /**
     * @param object
     * @param picId
     * @param excelParams
     * @param titleString
     * @param pictures
     * @param params
     * @throws Exception
     */
    private void saveImage(Object object, String picId, Map<String, ExcelImportEntity> excelParams,
                           String titleString, Map<String, PictureData> pictures,
                           ImportParams params) throws Exception {
        if (pictures == null) {
            return;
        }
        PictureData image = pictures.get(picId);
        if (image == null) {
            return;
        }
        byte[] data = image.getData();
        String fileName = "pic" + Math.round(Math.random() * 100000000000L);
        fileName += "." + PoiPublicUtil.getFileExtendName(data);
        if (excelParams.get(titleString).getSaveType() == 1) {
            String path = getSaveUrl(excelParams.get(titleString), object);
            File savefile = new File(path);
            if (!savefile.exists()) {
                savefile.mkdirs();
            }
            savefile = new File(path + File.separator + fileName);
            FileOutputStream fos = new FileOutputStream(savefile);
            try {
                fos.write(data);
            } finally {
                IOUtils.closeQuietly(fos);
            }
            setValues(excelParams.get(titleString), object,
                    getSaveUrl(excelParams.get(titleString), object) + File.separator + fileName);
        } else {
            setValues(excelParams.get(titleString), object, data);
        }
    }

    private void createErrorCellStyle(Workbook workbook) {
        errorCellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setColor(Font.COLOR_RED);
        errorCellStyle.setFont(font);
    }

}
