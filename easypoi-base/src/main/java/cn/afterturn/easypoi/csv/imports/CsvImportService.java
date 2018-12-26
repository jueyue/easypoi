package cn.afterturn.easypoi.csv.imports;

import cn.afterturn.easypoi.csv.entity.CsvImportParams;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import cn.afterturn.easypoi.excel.entity.params.ExcelCollectionParams;
import cn.afterturn.easypoi.excel.entity.params.ExcelImportEntity;
import cn.afterturn.easypoi.excel.entity.result.ExcelVerifyHandlerResult;
import cn.afterturn.easypoi.excel.imports.CellValueService;
import cn.afterturn.easypoi.excel.imports.base.ImportBaseService;
import cn.afterturn.easypoi.exception.excel.ExcelImportException;
import cn.afterturn.easypoi.exception.excel.enums.ExcelImportEnum;
import cn.afterturn.easypoi.handler.inter.IExcelModel;
import cn.afterturn.easypoi.handler.inter.IReadHandler;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import cn.afterturn.easypoi.util.PoiReflectorUtil;
import cn.afterturn.easypoi.util.PoiValidationUtil;
import cn.afterturn.easypoi.util.UnicodeInputStream;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.apache.poi.ss.usermodel.Cell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.util.*;

/**
 * Csv 导入服务
 * @author by jueyue on 18-10-3.
 */
public class CsvImportService extends ImportBaseService {

    private final static Logger LOGGER = LoggerFactory.getLogger(CsvImportService.class);

    private CellValueService cellValueServer;

    private boolean verifyFail = false;

    public CsvImportService() {
        this.cellValueServer = new CellValueService();
    }

    public <T> List<T> readExcel(InputStream inputstream, Class<?> pojoClass, CsvImportParams params, IReadHandler readHandler) {
        List collection = new ArrayList();
        try {
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

            inputstream = new PushbackInputStream(inputstream, 3);
            byte[] head = new byte[3];
            inputstream.read(head);
            // 判断 UTF8 是不是有 BOM
            if (head[0] == -17 && head[1] == -69 && head[2] == -65) {
                ((PushbackInputStream) inputstream).unread(head, 0, 3);
                inputstream = new UnicodeInputStream(inputstream);
            } else {
                ((PushbackInputStream) inputstream).unread(head, 0, 3);
            }
            BufferedReader rows = new BufferedReader(new InputStreamReader(inputstream, params.getEncoding()));
            for (int j = 0; j < params.getTitleRows(); j++) {
                rows.readLine();
            }
            Map<Integer, String> titlemap = getTitleMap(rows, params, excelCollection, excelParams);
            int readRow = 0;
            //跳过无效行
            for (int i = 0; i < params.getStartRows(); i++) {
                rows.readLine();
            }
            //判断index 和集合,集合情况默认为第一列
            if (excelCollection.size() > 0 && params.getKeyIndex() == null) {
                params.setKeyIndex(0);
            }
            StringBuilder errorMsg;
            String row = null;
            Object object = null;
            String[] cells;
            while ((row = rows.readLine()) != null) {
                if (StringUtils.isEmpty(row)) {
                    continue;
                }
                errorMsg = new StringBuilder();
                cells = row.split(params.getSpiltMark(), -1);
                // 判断是集合元素还是不是集合元素,如果是就继续加入这个集合,不是就创建新的对象
                // keyIndex 如果为空就不处理,仍然处理这一行
                if (params.getKeyIndex() != null && (cells[params.getKeyIndex()] == null
                        || StringUtils.isEmpty(cells[params.getKeyIndex()]))
                        && object != null) {
                    for (ExcelCollectionParams param : excelCollection) {
                        addListContinue(object, param, row, titlemap, targetId, params, errorMsg);
                    }
                } else {
                    object = PoiPublicUtil.createObject(pojoClass, targetId);
                    try {
                        Set<Integer> keys = titlemap.keySet();
                        for (Integer cn : keys) {
                            String titleString = (String) titlemap.get(cn);
                            if (excelParams.containsKey(titleString) || Map.class.equals(pojoClass)) {
                                try {
                                    saveFieldValue(params, object, cells[cn], excelParams, titleString);
                                } catch (ExcelImportException e) {
                                    // 如果需要去校验就忽略,这个错误,继续执行
                                    if (params.isNeedVerify() && ExcelImportEnum.GET_VALUE_ERROR.equals(e.getType())) {
                                        errorMsg.append(" ").append(titleString).append(ExcelImportEnum.GET_VALUE_ERROR.getMsg());
                                    }
                                }
                            }
                        }
                        for (ExcelCollectionParams param : excelCollection) {
                            addListContinue(object, param, row, titlemap, targetId, params, errorMsg);
                        }
                        if (verifyingDataValidity(object, params, pojoClass, errorMsg)) {
                            if (readHandler != null) {
                                readHandler.handler(object);
                            } else {
                                collection.add(object);
                            }
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
            if (readHandler != null) {
                readHandler.doAfterAll();
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }
        return collection;
    }

    private void addListContinue(Object object, ExcelCollectionParams param, String row,
                                 Map<Integer, String> titlemap, String targetId,
                                 CsvImportParams params, StringBuilder errorMsg) throws Exception {
        Collection collection = (Collection) PoiReflectorUtil.fromCache(object.getClass())
                .getValue(object, param.getName());
        Object entity = PoiPublicUtil.createObject(param.getType(), targetId);
        // 是否需要加上这个对象
        boolean isUsed = false;
        String[] cells = row.split(params.getSpiltMark());
        for (int i = 0; i < cells.length; i++) {
            String cell = cells[i];
            String titleString = (String) titlemap.get(i);
            if (param.getExcelParams().containsKey(titleString)) {
                try {
                    saveFieldValue(params, entity, cell, param.getExcelParams(), titleString);
                } catch (ExcelImportException e) {
                    // 如果需要去校验就忽略,这个错误,继续执行
                    if (params.isNeedVerify() && ExcelImportEnum.GET_VALUE_ERROR.equals(e.getType())) {
                        errorMsg.append(" ").append(titleString).append(ExcelImportEnum.GET_VALUE_ERROR.getMsg());
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
     * 校验数据合法性
     */
    private boolean verifyingDataValidity(Object object, CsvImportParams params,
                                          Class<?> pojoClass, StringBuilder fieldErrorMsg) {
        boolean isAdd = true;
        Cell cell = null;
        if (params.isNeedVerify()) {
            String errorMsg = PoiValidationUtil.validation(object, params.getVerifyGroup());
            if (StringUtils.isNotEmpty(errorMsg)) {
                if (object instanceof IExcelModel) {
                    IExcelModel model = (IExcelModel) object;
                    model.setErrorMsg(errorMsg);
                }
                isAdd = false;
                verifyFail = true;
            }
        }
        if (params.getVerifyHandler() != null) {
            ExcelVerifyHandlerResult result = params.getVerifyHandler().verifyHandler(object);
            if (!result.isSuccess()) {
                if (object instanceof IExcelModel) {
                    IExcelModel model = (IExcelModel) object;
                    model.setErrorMsg((StringUtils.isNoneBlank(model.getErrorMsg())
                            ? model.getErrorMsg() + "," : "") + result.getMsg());
                }
                isAdd = false;
                verifyFail = true;
            }
        }
        if ((params.isNeedVerify() || params.getVerifyHandler() != null) && fieldErrorMsg.length() > 0) {
            if (object instanceof IExcelModel) {
                IExcelModel model = (IExcelModel) object;
                model.setErrorMsg((StringUtils.isNoneBlank(model.getErrorMsg())
                        ? model.getErrorMsg() + "," : "") + fieldErrorMsg.toString());
            }
            isAdd = false;
            verifyFail = true;
        }
        return isAdd;
    }

    /**
     * 保存字段值(获取值,校验值,追加错误信息)
     */
    private void saveFieldValue(CsvImportParams params, Object object, String cell,
                                Map<String, ExcelImportEntity> excelParams, String titleString) throws Exception {
        if (cell.startsWith(params.getTextMark()) && cell.endsWith(params.getTextMark())) {
            cell = cell.replaceFirst(cell, params.getTextMark());
            cell = cell.substring(0, cell.lastIndexOf(params.getTextMark()));
        }
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
     * 获取表格字段列名对应信息
     */
    private Map<Integer, String> getTitleMap(BufferedReader rows, CsvImportParams params,
                                             List<ExcelCollectionParams> excelCollection,
                                             Map<String, ExcelImportEntity> excelParams) throws IOException {
        Map<Integer, String> titlemap = new LinkedHashMap<Integer, String>();
        String collectionName = null;
        ExcelCollectionParams collectionParams = null;
        String row = null;
        String[] cellTitle;
        for (int j = 0; j < params.getHeadRows(); j++) {
            row = rows.readLine();
            if (row == null) {
                continue;
            }
            cellTitle = row.split(params.getSpiltMark());
            for (int i = 0; i < cellTitle.length; i++) {
                String value = cellTitle[i];
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

}
