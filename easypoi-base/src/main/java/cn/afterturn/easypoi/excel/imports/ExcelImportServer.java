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
package cn.afterturn.easypoi.excel.imports;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.params.ExcelCollectionParams;
import cn.afterturn.easypoi.excel.entity.params.ExcelImportEntity;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import cn.afterturn.easypoi.excel.entity.result.ExcelVerifyHanlderResult;
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
 * @author JueYue
 *  2014年6月26日 下午9:20:51
 */
@SuppressWarnings({ "rawtypes", "unchecked", "hiding" })
public class ExcelImportServer extends ImportBaseService {

    private final static Logger LOGGER     = LoggerFactory.getLogger(ExcelImportServer.class);

    private CellValueServer     cellValueServer;

    private boolean             verfiyFail = false;
    /**
     * 异常数据styler
     */
    private CellStyle           errorCellStyle;

    public ExcelImportServer() {
        this.cellValueServer = new CellValueServer();
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
                                 ImportParams params) throws Exception {
        Collection collection = (Collection) PoiReflectorUtil.fromCache(object.getClass())
            .getValue(object, param.getName());
        Object entity = PoiPublicUtil.createObject(param.getType(), targetId);
        String picId;
        boolean isUsed = false;// 是否需要加上这个对象
        for (int i = row.getFirstCellNum(); i < titlemap.size(); i++) {
            Cell cell = row.getCell(i);
            String titleString = (String) titlemap.get(i);
            if (param.getExcelParams().containsKey(titleString)) {
                if (param.getExcelParams().get(titleString).getType() == 2) {
                    picId = row.getRowNum() + "_" + i;
                    saveImage(object, picId, param.getExcelParams(), titleString, pictures, params);
                } else {
                    saveFieldValue(params, entity, cell, param.getExcelParams(), titleString, row);
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
     * @author JueYue
     *  2013-11-21
     * @param cell
     * @return
     */
    private String getKeyValue(Cell cell) {
        Object obj = PoiCellUtil.getCellValue(cell);
        return obj == null ? null : obj.toString().trim();
    }

    /**
     * 获取保存的真实路径
     * 
     * @param excelImportEntity
     * @param object
     * @return
     * @throws Exception
     */
    private String getSaveUrl(ExcelImportEntity excelImportEntity, Object object) throws Exception {
        String url = "";
        if (excelImportEntity.getSaveUrl().equals("upload")) {
            if (excelImportEntity.getMethods() != null
                && excelImportEntity.getMethods().size() > 0) {
                object = getFieldBySomeMethod(excelImportEntity.getMethods(), object);
            }
            url = object.getClass().getName()
                .split("\\.")[object.getClass().getName().split("\\.").length - 1];
            return excelImportEntity.getSaveUrl() + "/" + url;
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
        if (!Map.class.equals(pojoClass)) {
            Field fileds[] = PoiPublicUtil.getClassFields(pojoClass);
            ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
            if (etarget != null) {
                targetId = etarget.value();
            }
            getAllExcelField(targetId, fileds, excelParams, excelCollection, pojoClass, null);
        }
        Iterator<Row> rows = sheet.rowIterator();
        for (int j = 0; j < params.getTitleRows(); j++) {
            rows.next();
        }
        Map<Integer, String> titlemap = getTitleMap(rows, params, excelCollection);
        checkIsValidTemplate(titlemap, excelParams, params, excelCollection);
        Row row = null;
        Object object = null;
        String picId;
        int readRow = 0;
        while (rows.hasNext()
               && (row == null
                   || sheet.getLastRowNum() - row.getRowNum() > params.getLastOfInvalidRow())) {
            if (params.getReadRows() > 0 && readRow > params.getReadRows()) {
                break;
            }
            row = rows.next();
            // 判断是集合元素还是不是集合元素,如果是就继续加入这个集合,不是就创建新的对象
            // keyIndex 如果为空就不处理,仍然处理这一行
            if (params.getKeyIndex() != null
                && (row.getCell(params.getKeyIndex()) == null
                    || StringUtils.isEmpty(getKeyValue(row.getCell(params.getKeyIndex()))))
                && object != null) {
                for (ExcelCollectionParams param : excelCollection) {
                    addListContinue(object, param, row, titlemap, targetId, pictures, params);
                }
            } else {
                object = PoiPublicUtil.createObject(pojoClass, targetId);
                try {
                    for (int i = row.getFirstCellNum(), le = titlemap.size(); i < le; i++) {
                        Cell cell = row.getCell(i);
                        String titleString = (String) titlemap.get(i);
                        if (excelParams.containsKey(titleString) || Map.class.equals(pojoClass)) {
                            if (excelParams.get(titleString) != null
                                && excelParams.get(titleString).getType() == 2) {
                                picId = row.getRowNum() + "_" + i;
                                saveImage(object, picId, excelParams, titleString, pictures,
                                    params);
                            } else {
                                saveFieldValue(params, object, cell, excelParams, titleString, row);
                            }
                        }
                    }

                    for (ExcelCollectionParams param : excelCollection) {
                        addListContinue(object, param, row, titlemap, targetId, pictures, params);
                    }
                    if (verifyingDataValidity(object, row, params, pojoClass)) {
                        collection.add(object);
                    }
                } catch (ExcelImportException e) {
                    LOGGER.error("excel import error , row num:{},obj:{}",readRow, ReflectionToStringBuilder.toString(object));
                    if (!e.getType().equals(ExcelImportEnum.VERIFY_ERROR)) {
                        throw new ExcelImportException(e.getType(), e);
                    }
                } catch (Exception e) {
                    LOGGER.error("excel import error , row num:{},obj:{}",readRow, ReflectionToStringBuilder.toString(object));
                    throw new RuntimeException(e);
                }
            }
            readRow++;
        }
        return collection;
    }

    /**
     * 校验数据合法性
     * @param object
     * @param row
     * @param params
     * @param pojoClass
     * @return
     */
    private boolean verifyingDataValidity(Object object, Row row, ImportParams params,
                                          Class<?> pojoClass) {
        boolean isAdd = true;
        Cell cell = null;
        if (params.isNeedVerfiy()) {
            String errorMsg = PoiValidationUtil.validation(object);
            if (StringUtils.isNotEmpty(errorMsg)) {
                cell = row.createCell(row.getLastCellNum());
                cell.setCellValue(errorMsg);
                if (object instanceof IExcelModel) {
                    IExcelModel model = (IExcelModel) object;
                    model.setErrorMsg(errorMsg);
                } else {
                    isAdd = false;
                }
                verfiyFail = true;
            }
        }
        if (params.getVerifyHanlder() != null) {
            ExcelVerifyHanlderResult result = params.getVerifyHanlder().verifyHandler(object);
            if (!result.isSuccess()) {
                if (cell == null)
                    cell = row.createCell(row.getLastCellNum());
                cell.setCellValue((StringUtils.isNoneBlank(cell.getStringCellValue())
                    ? cell.getStringCellValue() + "," : "") + result.getMsg());
                if (object instanceof IExcelModel) {
                    IExcelModel model = (IExcelModel) object;
                    model.setErrorMsg((StringUtils.isNoneBlank(model.getErrorMsg())
                        ? model.getErrorMsg() + "," : "") + result.getMsg());
                } else {
                    isAdd = false;
                }
                verfiyFail = true;
            }
        }
        if (cell != null)
            cell.setCellStyle(errorCellStyle);
        return isAdd;
    }

    /**
     * 获取表格字段列名对应信息
     * @param rows
     * @param params
     * @param excelCollection
     * @return
     */
    private Map<Integer, String> getTitleMap(Iterator<Row> rows, ImportParams params,
                                             List<ExcelCollectionParams> excelCollection) {
        Map<Integer, String> titlemap = new HashMap<Integer, String>();
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
        return titlemap;
    }

    /**
     * 获取这个名称对应的集合信息
     * @param excelCollection 
     * @param collectionName
     * @return
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
     * 
     * @param inputstream
     * @param pojoClass
     * @param params
     * @return
     * @throws Exception
     */
    public ExcelImportResult importExcelByIs(InputStream inputstream, Class<?> pojoClass,
                                             ImportParams params) throws Exception {
        if (LOGGER.isDebugEnabled()) {
            LOGGER.debug("Excel import start ,class is {}", pojoClass);
        }
        List<T> result = new ArrayList<T>();
        Workbook book = null;
        boolean isXSSFWorkbook = true;
        if (!(inputstream.markSupported())) {
            inputstream = new PushbackInputStream(inputstream, 8);
        }
        if (POIFSFileSystem.hasPOIFSHeader(inputstream)) {
            book = new HSSFWorkbook(inputstream);
            isXSSFWorkbook = false;
        } else if (DocumentFactoryHelper.hasOOXMLHeader(inputstream)) {
            book = new XSSFWorkbook(OPCPackage.open(inputstream));
        }
        ExcelImportResult importResult = new ExcelImportResult(result, verfiyFail, book);
        createErrorCellStyle(book);
        Map<String, PictureData> pictures;
        for (int i = params.getStartSheetIndex(); i < params.getStartSheetIndex()
                                                      + params.getSheetNum(); i++) {
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug(" start to read excel by is ,startTime is {}", new Date().getTime());
            }
            if (isXSSFWorkbook) {
                pictures = PoiPublicUtil.getSheetPictrues07((XSSFSheet) book.getSheetAt(i),
                    (XSSFWorkbook) book);
            } else {
                pictures = PoiPublicUtil.getSheetPictrues03((HSSFSheet) book.getSheetAt(i),
                    (HSSFWorkbook) book);
            }
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug(" end to read excel by is ,endTime is {}", new Date().getTime());
            }
            result.addAll(importExcel(result, book.getSheetAt(i), pojoClass, params, pictures));
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug(" end to read excel list by pos ,endTime is {}", new Date().getTime());
            }
            if(params.isReadSingleCell()){
                readSingleCell(importResult, book.getSheetAt(i),params);
                if (LOGGER.isDebugEnabled()) {
                    LOGGER.debug(" read Key-Value ,endTime is {}", new Date().getTime());
                }
            }
        }
        if (params.isNeedSave()) {
            saveThisExcel(params, pojoClass, isXSSFWorkbook, book);
        }
        importResult.setVerfiyFail(verfiyFail);
        return importResult;
    }

    /**
     * 按照键值对的方式取得Excel里面的数据
     * @param result
     * @param sheet
     * @param params
     */
    private void readSingleCell(ExcelImportResult result, Sheet sheet, ImportParams params) {
        if(result.getMap() == null){
            result.setMap(new HashMap<String, Object>());
        }
        for (int i = 0; i < params.getTitleRows() + params.getHeadRows() + params.getStartRows(); i++) {
           getSingleCellValueForRow(result,sheet.getRow(i),params);
        }

        for (int i = sheet.getLastRowNum() - params.getLastOfInvalidRow(); i < sheet.getLastRowNum(); i++) {
            getSingleCellValueForRow(result,sheet.getRow(i),params);

        }
    }

    private void getSingleCellValueForRow(ExcelImportResult result, Row row, ImportParams params) {
        for (int j = row.getFirstCellNum(), le = row.getLastCellNum(); j < le; j++) {
            String text = PoiCellUtil.getCellValue(row.getCell(j));
            if(StringUtils.isNoneBlank(text) && text.endsWith(params.getKeyMark())){
                if(result.getMap().containsKey(text)){
                    if(result.getMap().get(text) instanceof String){
                        List<String> list = new ArrayList<String>();
                        list.add((String)result.getMap().get(text));
                        result.getMap().put(text,list);
                    }
                    ((List)result.getMap().get(text)).add(PoiCellUtil.getCellValue(row.getCell(++j)));
                }else{
                    result.getMap().put(text,PoiCellUtil.getCellValue(row.getCell(++j)));
                }

            }

        }
    }

    /**
     * 检查是不是合法的模板
     * @param titlemap
     * @param excelParams
     * @param params
     * @param excelCollection 
     */
    private void checkIsValidTemplate(Map<Integer, String> titlemap,
                                      Map<String, ExcelImportEntity> excelParams,
                                      ImportParams params,
                                      List<ExcelCollectionParams> excelCollection) {

        if (params.getImportFields() != null) {
            for (int i = 0, le = params.getImportFields().length; i < le; i++) {
                if (!titlemap.containsValue(params.getImportFields()[i])) {
                    throw new ExcelImportException(ExcelImportEnum.IS_NOT_A_VALID_TEMPLATE);
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
     * 
     * @param params
     * @param object
     * @param cell
     * @param excelParams
     * @param titleString
     * @param row
     * @throws Exception
     */
    private void saveFieldValue(ImportParams params, Object object, Cell cell,
                                Map<String, ExcelImportEntity> excelParams, String titleString,
                                Row row) throws Exception {
        Object value = cellValueServer.getValue(params.getDataHanlder(), object, cell, excelParams,
            titleString);
        if (object instanceof Map) {
            if (params.getDataHanlder() != null) {
                params.getDataHanlder().setMapValue((Map) object, titleString, value);
            } else {
                ((Map) object).put(titleString, value);
            }
        } else {
            setValues(excelParams.get(titleString), object, value);
        }
    }

    /**
     * 
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
            String path = PoiPublicUtil
                .getWebRootPath(getSaveUrl(excelParams.get(titleString), object));
            File savefile = new File(path);
            if (!savefile.exists()) {
                savefile.mkdirs();
            }
            savefile = new File(path + "/" + fileName);
            FileOutputStream fos = new FileOutputStream(savefile);
            try {
                fos.write(data);
            } finally {
                IOUtils.closeQuietly(fos);
            }
            setValues(excelParams.get(titleString), object,
                getSaveUrl(excelParams.get(titleString), object) + "/" + fileName);
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
