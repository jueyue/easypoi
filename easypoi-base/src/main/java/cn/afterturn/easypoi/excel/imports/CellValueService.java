/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package cn.afterturn.easypoi.excel.imports;

import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.sql.Time;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;

import cn.afterturn.easypoi.handler.inter.IExcelDictHandler;
import cn.afterturn.easypoi.util.PoiReflectorUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import cn.afterturn.easypoi.excel.entity.params.ExcelImportEntity;
import cn.afterturn.easypoi.excel.entity.sax.SaxReadCellEntity;
import cn.afterturn.easypoi.exception.excel.ExcelImportException;
import cn.afterturn.easypoi.exception.excel.enums.ExcelImportEnum;
import cn.afterturn.easypoi.handler.inter.IExcelDataHandler;
import cn.afterturn.easypoi.util.PoiPublicUtil;

/**
 * Cell 取值服务
 * 判断类型处理数据 1.判断Excel中的类型 2.根据replace替换值 3.handler处理数据 4.判断返回类型转化数据返回
 *
 * @author JueYue
 *  2014年6月26日 下午10:42:28
 */
public class CellValueService {

    private static final Logger LOGGER = LoggerFactory.getLogger(CellValueService.class);

    private List<String> handlerList = null;

    /**
     * 获取单元格内的值
     *
     * @param cell
     * @param entity
     * @return
     */
    private Object getCellValue(String xclass, Cell cell, ExcelImportEntity entity) {
        if (cell == null) {
            return "";
        }
        Object result = null;
        if ("class java.util.Date".equals(xclass) || "class java.sql.Date".equals(xclass)
                || ("class java.sql.Time").equals(xclass)
                || ("class java.sql.Timestamp").equals(xclass)) {
            /*
            if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                // 日期格式
                result = cell.getDateCellValue();
            } else {
                cell.setCellType(Cell.CELL_TYPE_STRING);
                result = getDateData(entity, cell.getStringCellValue());
            }*/
            //FIX: 单元格yyyyMMdd数字时候使用 cell.getDateCellValue() 解析出的日期错误
            if (Cell.CELL_TYPE_NUMERIC == cell.getCellType() && DateUtil.isCellDateFormatted(cell)) {
                result = DateUtil.getJavaDate(cell.getNumericCellValue());
            } else {
                cell.setCellType(CellType.STRING);
                result = getDateData(entity, cell.getStringCellValue());
            }
            if (("class java.sql.Date").equals(xclass)) {
                result = new java.sql.Date(((Date) result).getTime());
            }
            if (("class java.sql.Time").equals(xclass)) {
                result = new Time(((Date) result).getTime());
            }
            if (("class java.sql.Timestamp").equals(xclass)) {
                result = new Timestamp(((Date) result).getTime());
            }
        } else if (Cell.CELL_TYPE_NUMERIC == cell.getCellType() && DateUtil.isCellDateFormatted(cell)) {
            result = DateUtil.getJavaDate(cell.getNumericCellValue());
        } else {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    result = cell.getRichStringCellValue() == null ? ""
                            : cell.getRichStringCellValue().getString();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        if ("class java.lang.String".equals(xclass)) {
                            result = formateDate(entity, cell.getDateCellValue());
                        }
                    } else {
                        result = readNumericCell(cell);
                    }
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    result = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_BLANK:
                    break;
                case Cell.CELL_TYPE_ERROR:
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    try {
                        result = readNumericCell(cell);
                    } catch (Exception e1) {
                        try {
                            result = cell.getRichStringCellValue() == null ? ""
                                    : cell.getRichStringCellValue().getString();
                        } catch (Exception e2) {
                            throw new RuntimeException("获取公式类型的单元格失败", e2);
                        }
                    }
                    break;
                default:
                    break;
            }
        }
        return result;
    }

    private Object readNumericCell(Cell cell) {
        Object result = null;
        double value = cell.getNumericCellValue();
        if (((int) value) == value) {
            result = (int) value;
        } else {
            result = value;
        }
        return result;
    }

    /**
     * 获取日期类型数据
     *
     * @author JueYue
     *  2013年11月26日
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

    private String formateDate(ExcelImportEntity entity, Date value) {
        if (StringUtils.isNotEmpty(entity.getFormat()) && value != null) {
            SimpleDateFormat format = new SimpleDateFormat(entity.getFormat());
            return format.format(value);
        }
        return null;
    }

    /**
     * 获取cell的值
     *  @param object
     * @param cell
     * @param excelParams
     * @param titleString
     * @param dictHandler
     */
    public Object getValue(IExcelDataHandler<?> dataHandler, Object object, Cell cell,
                           Map<String, ExcelImportEntity> excelParams,
                           String titleString, IExcelDictHandler dictHandler) throws Exception {
        ExcelImportEntity entity = excelParams.get(titleString);
        String xclass = "class java.lang.Object";
        Class clazz = null;
        if (!(object instanceof Map)) {
            Method setMethod = entity.getMethods() != null && entity.getMethods().size() > 0
                    ? entity.getMethods().get(entity.getMethods().size() - 1) : entity.getMethod();
            Type[] ts = setMethod.getGenericParameterTypes();
            xclass = ts[0].toString();
            clazz = (Class) ts[0];
        }
        Object result = getCellValue(xclass, cell, entity);
        if (entity != null) {
            result = hanlderSuffix(entity.getSuffix(), result);
            result = replaceValue(entity.getReplace(), result);
            result = replaceValue(entity.getReplace(), result);
            if(dictHandler != null && StringUtils.isNoneBlank(entity.getDict())){
                dictHandler.toValue(entity.getDict(), object, entity.getName(), result);
            }
        }
        result = handlerValue(dataHandler, object, result, titleString);
        return getValueByType(xclass, result, entity, clazz);
    }

    /**
     * 获取cell值
     * @param dataHandler
     * @param object
     * @param cellEntity
     * @param excelParams
     * @param titleString
     * @return
     */
    public Object getValue(IExcelDataHandler<?> dataHandler, Object object,
                           SaxReadCellEntity cellEntity, Map<String, ExcelImportEntity> excelParams,
                           String titleString) {
        ExcelImportEntity entity = excelParams.get(titleString);
        Method setMethod = entity.getMethods() != null && entity.getMethods().size() > 0
                ? entity.getMethods().get(entity.getMethods().size() - 1) : entity.getMethod();
        Type[] ts = setMethod.getGenericParameterTypes();
        String xclass = ts[0].toString();
        Object result = cellEntity.getValue();
        result = hanlderSuffix(entity.getSuffix(), result);
        result = replaceValue(entity.getReplace(), result);
        result = handlerValue(dataHandler, object, result, titleString);
        return getValueByType(xclass, result, entity, (Class) ts[0]);
    }

    /**
     * 把后缀删除掉
     * @param result
     * @param suffix
     * @return
     */
    private Object hanlderSuffix(String suffix, Object result) {
        if (StringUtils.isNotEmpty(suffix) && result != null
                && result.toString().endsWith(suffix)) {
            String temp = result.toString();
            return temp.substring(0, temp.length() - suffix.length());
        }
        return result;
    }

    /**
     * 根据返回类型获取返回值
     *
     * @param xclass
     * @param result
     * @param entity
     * @param clazz
     * @return
     */
    private Object getValueByType(String xclass, Object result, ExcelImportEntity entity, Class clazz) {
        try {
            //过滤空和空字符串,如果基本类型null会在上层抛出,这里就不处理了
            if (result == null || StringUtils.isBlank(result.toString())) {
                return null;
            }
            if ("class java.util.Date".equals(xclass)) {
                return result;
            }
            if ("class java.lang.Boolean".equals(xclass) || "boolean".equals(xclass)) {
                return Boolean.valueOf(String.valueOf(result));
            }
            if ("class java.lang.Double".equals(xclass) || "double".equals(xclass)) {
                return Double.valueOf(String.valueOf(result));
            }
            if ("class java.lang.Long".equals(xclass) || "long".equals(xclass)) {
                try {
                    return Long.valueOf(String.valueOf(result));
                } catch (Exception e) {
                    //格式错误的时候,就用double,然后获取Int值
                    return Double.valueOf(String.valueOf(result)).longValue();
                }
            }
            if ("class java.lang.Float".equals(xclass) || "float".equals(xclass)) {
                return Float.valueOf(String.valueOf(result));
            }
            if ("class java.lang.Integer".equals(xclass) || "int".equals(xclass)) {
                try {
                    return Integer.valueOf(String.valueOf(result));
                } catch (Exception e) {
                    //格式错误的时候,就用double,然后获取Int值
                    return Double.valueOf(String.valueOf(result)).intValue();
                }
            }
            if ("class java.math.BigDecimal".equals(xclass)) {
                return new BigDecimal(String.valueOf(result));
            }
            if ("class java.lang.String".equals(xclass)) {
                //针对String 类型,但是Excel获取的数据却不是String,比如Double类型,防止科学计数法
                if (result instanceof String) {
                    return result;
                }
                // double类型防止科学计数法
                if (result instanceof Double) {
                    return PoiPublicUtil.doubleToString((Double) result);
                }
                return String.valueOf(result);
            }
            if (clazz.isEnum()) {
                if (StringUtils.isNotEmpty(entity.getEnumImportMethod())) {
                    return PoiReflectorUtil.fromCache(clazz).execEnumStaticMethod(entity.getEnumImportMethod(),result);
                } else {
                    return Enum.valueOf(clazz,result.toString());
                }
            }
            return result;
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelImportException(ExcelImportEnum.GET_VALUE_ERROR);
        }
    }

    /**
     * 调用处理接口处理值
     *
     * @param dataHandler
     * @param object
     * @param result
     * @param titleString
     * @return
     */
    @SuppressWarnings({"unchecked", "rawtypes"})
    private Object handlerValue(IExcelDataHandler dataHandler, Object object, Object result,
                                String titleString) {
        if (dataHandler == null || dataHandler.getNeedHandlerFields() == null
                || dataHandler.getNeedHandlerFields().length == 0) {
            return result;
        }
        if (handlerList == null) {
            handlerList = Arrays.asList(dataHandler.getNeedHandlerFields());
        }
        if (handlerList.contains(titleString)) {
            return dataHandler.importHandler(object, titleString, result);
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