package org.jeecgframework.poi.excel.imports;

import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.jeecgframework.poi.excel.entity.params.ExcelImportEntity;
import org.jeecgframework.poi.excel.entity.sax.SaxReadCellEntity;
import org.jeecgframework.poi.exception.excel.ExcelImportException;
import org.jeecgframework.poi.exception.excel.enums.ExcelImportEnum;
import org.jeecgframework.poi.handler.inter.IExcelDataHandler;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Cell 取值服务
 * 判断类型处理数据 1.判断Excel中的类型 2.根据replace替换值 3.handler处理数据 4.判断返回类型转化数据返回
 * 
 * @author JueYue
 * @date 2014年6月26日 下午10:42:28
 */
public class CellValueServer {

    private static final Logger LOGGER      = LoggerFactory.getLogger(CellValueServer.class);

    private List<String>        hanlderList = null;

    /**
     * 获取单元格内的值
     * 
     * @param xclass
     * @param cell
     * @param entity
     * @return
     */
    private Object getCellValue(String xclass, Cell cell, ExcelImportEntity entity) {
        if (cell == null) {
            return "";
        }
        Object result = null;
        // 日期格式比较特殊,和cell格式不一致
        if (xclass.equals("class java.util.Date")) {
            if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                // 日期格式
                result = cell.getDateCellValue();
            } else {
                cell.setCellType(Cell.CELL_TYPE_STRING);
                result = getDateData(entity, cell.getStringCellValue());
            }
        } else if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
            result = cell.getNumericCellValue();
        } else if (Cell.CELL_TYPE_BOOLEAN == cell.getCellType()) {
            result = cell.getBooleanCellValue();
        } else {
            result = cell.getStringCellValue();
        }
        return result;
    }

    /**
     * 获取日期类型数据
     * 
     * @Author JueYue
     * @date 2013年11月26日
     * @param entity
     * @param value
     * @return
     */
    private Date getDateData(ExcelImportEntity entity, String value) {
        if (StringUtils.isNotEmpty(entity.getFormat()) && StringUtils.isNotEmpty(value)) {
            SimpleDateFormat format = new SimpleDateFormat(entity.getFormat());
            try {
                return format.parse(value);
            } catch (ParseException e) {
                LOGGER.error("时间格式化失败,格式化:{},值:{}", entity.getFormat(), value);
                throw new ExcelImportException(ExcelImportEnum.GET_VALUE_ERROR);
            }
        }
        return null;
    }

    /**
     * 获取cell的值
     * 
     * @param object
     * @param excelParams
     * @param cell
     * @param titleString
     */
    public Object getValue(IExcelDataHandler dataHanlder, Object object, Cell cell,
                           Map<String, ExcelImportEntity> excelParams, String titleString)
                                                                                          throws Exception {
        ExcelImportEntity entity = excelParams.get(titleString);
        Method setMethod = entity.getMethods() != null && entity.getMethods().size() > 0 ? entity
            .getMethods().get(entity.getMethods().size() - 1) : entity.getMethod();
        Type[] ts = setMethod.getGenericParameterTypes();
        String xclass = ts[0].toString();
        Object result = getCellValue(xclass, cell, entity);
        result = replaceValue(entity.getReplace(), result);
        result = hanlderValue(dataHanlder, object, result, titleString);
        return getValueByType(xclass, result);
    }

    /**
     * 获取cell值
     * @param dataHanlder
     * @param object
     * @param entity
     * @param excelParams
     * @param titleString
     * @return
     */
    public Object getValue(IExcelDataHandler dataHanlder, Object object,
                           SaxReadCellEntity cellEntity,
                           Map<String, ExcelImportEntity> excelParams, String titleString) {
        ExcelImportEntity entity = excelParams.get(titleString);
        Method setMethod = entity.getMethods() != null && entity.getMethods().size() > 0 ? entity
            .getMethods().get(entity.getMethods().size() - 1) : entity.getMethod();
        Type[] ts = setMethod.getGenericParameterTypes();
        String xclass = ts[0].toString();
        Object result = cellEntity.getValue();
        result = replaceValue(entity.getReplace(), result);
        result = hanlderValue(dataHanlder, object, result, titleString);
        return getValueByType(xclass, result);
    }

    /**
     * 根据返回类型获取返回值
     * 
     * @param xclass
     * @param result
     * @return
     */
    private Object getValueByType(String xclass, Object result) {
        try {
            if (xclass.equals("class java.util.Date")) {
                return result;
            }
            if (xclass.equals("class java.lang.Boolean") || xclass.equals("boolean")) {
                return Boolean.valueOf(String.valueOf(result));
            }
            if (xclass.equals("class java.lang.Double") || xclass.equals("double")) {
                return Double.valueOf(String.valueOf(result));
            }
            if (xclass.equals("class java.lang.Long") || xclass.equals("long")) {
                return Long.valueOf(String.valueOf(result));
            }
            if (xclass.equals("class java.lang.Integer") || xclass.equals("int")) {
                return Integer.valueOf(String.valueOf(result));
            }
            if (xclass.equals("class java.math.BigDecimal")) {
                return new BigDecimal(String.valueOf(result));
            }
            if (xclass.equals("class java.lang.String")) {
                return String.valueOf(result);
            }
            return result;
        } catch (Exception e) {
            LOGGER.error(e.getMessage(),e);
            throw new ExcelImportException(ExcelImportEnum.GET_VALUE_ERROR);
        }
    }

    /**
     * 调用处理接口处理值
     * 
     * @param dataHanlder
     * @param object
     * @param result
     * @param titleString
     * @return
     */
    private Object hanlderValue(IExcelDataHandler dataHanlder, Object object, Object result,
                                String titleString) {
        if (dataHanlder == null) {
            return result;
        }
        if (hanlderList == null) {
            hanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
        }
        if (hanlderList.contains(titleString)) {
            return dataHanlder.importHandler(object, titleString, result);
        }
        return result;
    }

    /**
     * 替换值
     * 
     * @param replace
     * @param result
     * @return
     */
    private Object replaceValue(String[] replace, Object result) {
        if (replace != null && replace.length > 0) {
            String temp = String.valueOf(result);
            String[] tempArr;
            for (int i = 0; i < replace.length; i++) {
                tempArr = replace[i].split("_");
                if (temp.equals(tempArr[0])) {
                    return tempArr[1];
                }
            }
        }
        return result;
    }
}
