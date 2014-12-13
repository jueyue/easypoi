package org.jeecgframework.poi.excel.imports;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.annotation.ExcelVerify;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.jeecgframework.poi.excel.entity.params.ExcelCollectionParams;
import org.jeecgframework.poi.excel.entity.params.ExcelImportEntity;
import org.jeecgframework.poi.excel.entity.params.ExcelVerifyEntity;
import org.jeecgframework.poi.excel.entity.result.ExcelImportResult;
import org.jeecgframework.poi.excel.entity.result.ExcelVerifyHanlderResult;
import org.jeecgframework.poi.excel.imports.verifys.VerifyHandlerServer;
import org.jeecgframework.poi.util.POIPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Excel 导入服务
 * 
 * @author JueYue
 * @date 2014年6月26日 下午9:20:51
 */
@SuppressWarnings({ "rawtypes", "unchecked", "hiding" })
public class ExcelImportServer {

    private final static Logger LOGGER     = LoggerFactory.getLogger(ExcelImportServer.class);

    private CellValueServer     cellValueServer;

    private VerifyHandlerServer verifyHandlerServer;

    private boolean             verfiyFail = false;

    public ExcelImportServer() {
        this.cellValueServer = new CellValueServer();
        this.verifyHandlerServer = new VerifyHandlerServer();
    }

    /**
     * 把这个注解解析放到类型对象中
     * 
     * @param targetId
     * @param field
     * @param excelEntity
     * @param pojoClass
     * @param getMethods
     * @param temp
     * @throws Exception
     */
    private void addEntityToMap(String targetId, Field field, ExcelImportEntity excelEntity,
                                Class<?> pojoClass, List<Method> getMethods,
                                Map<String, ExcelImportEntity> temp) throws Exception {
        Excel excel = field.getAnnotation(Excel.class);
        excelEntity = new ExcelImportEntity();
        excelEntity.setType(excel.type());
        excelEntity.setSaveUrl(excel.savePath());
        excelEntity.setSaveType(excel.imageType());
        excelEntity.setReplace(excel.replace());
        excelEntity.setDatabaseFormat(excel.databaseFormat());
        excelEntity.setVerify(getImportVerify(field));
        getExcelField(targetId, field, excelEntity, excel, pojoClass);
        if (getMethods != null) {
            List<Method> newMethods = new ArrayList<Method>();
            newMethods.addAll(getMethods);
            newMethods.add(excelEntity.getMethod());
            excelEntity.setMethods(newMethods);
        }
        temp.put(excelEntity.getName(), excelEntity);

    }

    /***
     * 向List里面继续添加元素
     * 
     * @param exclusions
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
                                 Map<String, PictureData> pictures, ImportParams params)
                                                                                        throws Exception {
        Collection collection = (Collection) POIPublicUtil.getMethod(param.getName(),
            object.getClass()).invoke(object, new Object[] {});
        Object entity = POIPublicUtil.createObject(param.getType(), targetId);
        String picId;
        boolean isUsed = false;// 是否需要加上这个对象
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
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
     * 获取需要导出的全部字段
     * 
     * 
     * @param exclusions
     * @param targetId
     *            目标ID
     * @param fields
     * @param excelCollection
     * @throws Exception
     */
    private void getAllExcelField(String targetId, Field[] fields,
                                  Map<String, ExcelImportEntity> excelParams,
                                  List<ExcelCollectionParams> excelCollection, Class<?> pojoClass,
                                  List<Method> getMethods) throws Exception {
        ExcelImportEntity excelEntity = null;
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            if (POIPublicUtil.isNotUserExcelUserThis(null, field, targetId)) {
                continue;
            }
            if (POIPublicUtil.isCollection(field.getType())) {
                // 集合对象设置属性
                ExcelCollectionParams collection = new ExcelCollectionParams();
                collection.setName(field.getName());
                Map<String, ExcelImportEntity> temp = new HashMap<String, ExcelImportEntity>();
                ParameterizedType pt = (ParameterizedType) field.getGenericType();
                Class<?> clz = (Class<?>) pt.getActualTypeArguments()[0];
                collection.setType(clz);
                getExcelFieldList(targetId, POIPublicUtil.getClassFields(clz), clz, temp, null);
                collection.setExcelParams(temp);
                excelCollection.add(collection);
            } else if (POIPublicUtil.isJavaClass(field)) {
                addEntityToMap(targetId, field, excelEntity, pojoClass, getMethods, excelParams);
            } else {
                List<Method> newMethods = new ArrayList<Method>();
                if (getMethods != null) {
                    newMethods.addAll(getMethods);
                }
                newMethods.add(POIPublicUtil.getMethod(field.getName(), pojoClass));
                getAllExcelField(targetId, POIPublicUtil.getClassFields(field.getType()),
                    excelParams, excelCollection, field.getType(), newMethods);
            }
        }
    }

    private void getExcelField(String targetId, Field field, ExcelImportEntity excelEntity,
                               Excel excel, Class<?> pojoClass) throws Exception {
        excelEntity.setName(getExcelName(excel.name(), targetId));
        String fieldname = field.getName();
        excelEntity.setMethod(POIPublicUtil.getMethod(fieldname, pojoClass, field.getType()));
        if (StringUtils.isEmpty(excel.importFormat())) {
            excelEntity.setFormat(excel.format());
        } else {
            excelEntity.setFormat(excel.importFormat());
        }
    }

    private void getExcelFieldList(String targetId, Field[] fields, Class<?> pojoClass,
                                   Map<String, ExcelImportEntity> temp, List<Method> getMethods)
                                                                                                throws Exception {
        ExcelImportEntity excelEntity = null;
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            if (POIPublicUtil.isNotUserExcelUserThis(null, field, targetId)) {
                continue;
            }
            if (POIPublicUtil.isJavaClass(field)) {
                addEntityToMap(targetId, field, excelEntity, pojoClass, getMethods, temp);
            } else {
                List<Method> newMethods = new ArrayList<Method>();
                if (getMethods != null) {
                    newMethods.addAll(getMethods);
                }
                newMethods
                    .add(POIPublicUtil.getMethod(field.getName(), pojoClass, field.getType()));
                getExcelFieldList(targetId, POIPublicUtil.getClassFields(field.getType()),
                    field.getType(), temp, newMethods);
            }
        }
    }

    /**
     * 判断在这个单元格显示的名称
     * 
     * @param exportName
     * @param targetId
     * @return
     */
    private String getExcelName(String exportName, String targetId) {
        if (exportName.indexOf("_") < 0) {
            return exportName;
        }
        String[] arr = exportName.split(",");
        for (String str : arr) {
            if (str.indexOf(targetId) != -1) {
                return str.split("_")[0];
            }
        }
        return null;
    }

    private Object getFieldBySomeMethod(List<Method> list, Object t) throws Exception {
        Method m;
        for (int i = 0; i < list.size() - 1; i++) {
            m = list.get(i);
            t = m.invoke(t, new Object[] {});
        }
        return t;
    }

    /**
     * 获取导入校验参数
     * 
     * @param field
     * @return
     */
    private ExcelVerifyEntity getImportVerify(Field field) {
        ExcelVerify verify = field.getAnnotation(ExcelVerify.class);
        if (verify != null) {
            ExcelVerifyEntity entity = new ExcelVerifyEntity();
            entity.setEmail(verify.isEmail());
            entity.setInterHandler(verify.interHandler());
            entity.setMaxLength(verify.maxLength());
            entity.setMinLength(verify.minLength());
            entity.setMobile(verify.isMobile());
            entity.setNotNull(verify.notNull());
            entity.setRegex(verify.regex());
            entity.setRegexTip(verify.regexTip());
            entity.setTel(verify.isTel());
            return entity;
        }
        return null;
    }

    /**
     * 获取key的值,针对不同类型获取不同的值
     * 
     * @Author JueYue
     * @date 2013-11-21
     * @param cell
     * @return
     */
    private String getKeyValue(Cell cell) {
        Object obj = null;
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                obj = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                obj = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                obj = cell.getNumericCellValue();
                break;
        }
        return obj == null ? null : obj.toString();
    }

    /**
     * 获取保存的Excel 的真实路径
     * 
     * @param params
     * @param pojoClass
     * @return
     * @throws Exception
     */
    private String getSaveExcelUrl(ImportParams params, Class<?> pojoClass) throws Exception {
        String url = "";
        if (params.getSaveUrl().equals("upload/excelUpload")) {
            url = pojoClass.getName().split("\\.")[pojoClass.getName().split("\\.").length - 1];
            return params.getSaveUrl() + "/" + url;
        }
        return params.getSaveUrl();
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
            if (excelImportEntity.getMethods() != null && excelImportEntity.getMethods().size() > 0) {
                object = getFieldBySomeMethod(excelImportEntity.getMethods(), object);
            }
            url = object.getClass().getName().split("\\.")[object.getClass().getName().split("\\.").length - 1];
            return excelImportEntity.getSaveUrl() + "/"
                   + url.substring(0, url.lastIndexOf("Entity"));
        }
        return excelImportEntity.getSaveUrl();
    }

    private <T> List<T> importExcel(Collection<T> result, Sheet sheet, Class<?> pojoClass,
                                    ImportParams params, Map<String, PictureData> pictures)
                                                                                           throws Exception {
        List collection = new ArrayList();
        Map<String, ExcelImportEntity> excelParams = new HashMap<String, ExcelImportEntity>();
        List<ExcelCollectionParams> excelCollection = new ArrayList<ExcelCollectionParams>();
        Field fileds[] = POIPublicUtil.getClassFields(pojoClass);
        ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
        String targetId = null;
        if (etarget != null) {
            targetId = etarget.value();
        }
        getAllExcelField(targetId, fileds, excelParams, excelCollection, pojoClass, null);
        Iterator<Row> rows = sheet.rowIterator();
        for (int j = 0; j < params.getTitleRows(); j++) {
            rows.next();
        }
        Row row = null;
        Iterator<Cell> cellTitle;
        Map<Integer, String> titlemap = new HashMap<Integer, String>();
        for (int j = 0; j < params.getHeadRows(); j++) {
            row = rows.next();
            cellTitle = row.cellIterator();
            int i = row.getFirstCellNum();
            while (cellTitle.hasNext()) {
                Cell cell = cellTitle.next();
                String value = cell.getStringCellValue();
                if (!StringUtils.isEmpty(value)) {
                    titlemap.put(i, value);
                }
                i = i + 1;
            }
        }
        Object object = null;
        String picId;
        while (rows.hasNext()) {
            row = rows.next();
            // 判断是集合元素还是不是集合元素,如果是就继续加入这个集合,不是就创建新的对象
            if ((row.getCell(params.getKeyIndex()) == null || StringUtils.isEmpty(getKeyValue(row
                .getCell(params.getKeyIndex())))) && object != null) {
                for (ExcelCollectionParams param : excelCollection) {
                    addListContinue(object, param, row, titlemap, targetId, pictures, params);
                }
            } else {
                object = POIPublicUtil.createObject(pojoClass, targetId);
                for (int i = row.getFirstCellNum(), le = row.getLastCellNum(); i < le; i++) {
                    Cell cell = row.getCell(i);
                    String titleString = (String) titlemap.get(i);
                    if (excelParams.containsKey(titleString)) {
                        if (excelParams.get(titleString).getType() == 2) {
                            picId = row.getRowNum() + "_" + i;
                            saveImage(object, picId, excelParams, titleString, pictures, params);
                        } else {
                            saveFieldValue(params, object, cell, excelParams, titleString, row);
                        }
                    }
                }
                for (ExcelCollectionParams param : excelCollection) {
                    addListContinue(object, param, row, titlemap, targetId, pictures, params);
                }
                collection.add(object);
            }
        }
        return collection;
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
        } else if (POIXMLDocument.hasOOXMLHeader(inputstream)) {
            book = new XSSFWorkbook(OPCPackage.open(inputstream));
        }
        Map<String, PictureData> pictures;
        for (int i = 0; i < params.getSheetNum(); i++) {
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug(" start to read excel by is ,startTime is {}", new Date().getTime());
            }
            if (isXSSFWorkbook) {
                pictures = POIPublicUtil.getSheetPictrues07((XSSFSheet) book.getSheetAt(i),
                    (XSSFWorkbook) book);
            } else {
                pictures = POIPublicUtil.getSheetPictrues03((HSSFSheet) book.getSheetAt(i),
                    (HSSFWorkbook) book);
            }
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug(" end to read excel by is ,endTime is {}", new Date().getTime());
            }
            result.addAll(importExcel(result, book.getSheetAt(i), pojoClass, params, pictures));
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug(" end to read excel list by pos ,endTime is {}", new Date().getTime());
            }
        }
        if (params.isNeedSave()) {
            saveThisExcel(params, pojoClass, isXSSFWorkbook, book);
        }
        return new ExcelImportResult(result, verfiyFail, book);
    }

    /**
     * 保存字段值(获取值,校验值,追加错误信息)
     * 
     * @param params
     * @param object
     * @param cell
     * @param excelParams
     * @param titleString
     * @param excelParams2
     * @param row
     * @throws Exception
     */
    private void saveFieldValue(ImportParams params, Object object, Cell cell,
                                Map<String, ExcelImportEntity> excelParams, String titleString,
                                Row row) throws Exception {
        ExcelVerifyHanlderResult verifyResult;
        Object value = cellValueServer.getValue(params.getDataHanlder(), object, cell, excelParams,
            titleString);
        verifyResult = verifyHandlerServer.verifyData(object, value, titleString,
            excelParams.get(titleString).getVerify(), params.getVerifyHanlder());
        if (verifyResult.isSuccess()) {
            setValues(excelParams.get(titleString), object, value);
        } else {
            Cell errorCell = row.createCell(row.getLastCellNum());
            errorCell.setCellValue(verifyResult.getMsg());
            verfiyFail = true;
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
        byte[] data = image.getData();
        String fileName = "pic" + Math.round(Math.random() * 100000000000L);
        fileName += "." + POIPublicUtil.getFileExtendName(data);
        if (excelParams.get(titleString).getSaveType() == 1) {
            String path = POIPublicUtil.getWebRootPath(getSaveUrl(excelParams.get(titleString),
                object));
            File savefile = new File(path);
            if (!savefile.exists()) {
                savefile.mkdirs();
            }
            savefile = new File(path + "/" + fileName);
            FileOutputStream fos = new FileOutputStream(savefile);
            fos.write(data);
            fos.close();
            setValues(excelParams.get(titleString), object,
                getSaveUrl(excelParams.get(titleString), object) + "/" + fileName);
        } else {
            setValues(excelParams.get(titleString), object, data);
        }
    }

    private void saveThisExcel(ImportParams params, Class<?> pojoClass, boolean isXSSFWorkbook,
                               Workbook book) throws Exception {
        String path = POIPublicUtil.getWebRootPath(getSaveExcelUrl(params, pojoClass));
        File savefile = new File(path);
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        SimpleDateFormat format = new SimpleDateFormat("yyyMMddHHmmss");
        FileOutputStream fos = new FileOutputStream(path + "/" + format.format(new Date()) + "_"
                                                    + Math.round(Math.random() * 100000)
                                                    + (isXSSFWorkbook == true ? ".xlsx" : ".xls"));
        book.write(fos);
        fos.close();
    }

    /**
     * 多个get 最后再set
     * 
     * @param setMethods
     * @param object
     */
    private void setFieldBySomeMethod(List<Method> setMethods, Object object, Object value)
                                                                                           throws Exception {
        Object t = getFieldBySomeMethod(setMethods, object);
        setMethods.get(setMethods.size() - 1).invoke(t, value);
    }

    /**
     * 
     * @param entity
     * @param object
     * @param value
     * @throws Exception
     */
    private void setValues(ExcelImportEntity entity, Object object, Object value) throws Exception {
        if (entity.getMethods() != null) {
            setFieldBySomeMethod(entity.getMethods(), object, value);
        } else {
            entity.getMethod().invoke(object, value);
        }
    }

}
