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
package cn.afterturn.easypoi.excel.export.base;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelCollection;
import cn.afterturn.easypoi.excel.annotation.ExcelEntity;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.entity.vo.PoiBaseConstants;
import cn.afterturn.easypoi.exception.excel.ExcelExportException;
import cn.afterturn.easypoi.exception.excel.enums.ExcelExportEnum;
import cn.afterturn.easypoi.handler.inter.IExcelDataHandler;
import cn.afterturn.easypoi.handler.inter.IExcelDictHandler;
import cn.afterturn.easypoi.handler.inter.IExcelI18nHandler;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import cn.afterturn.easypoi.util.PoiReflectorUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.apache.commons.lang3.math.NumberUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.*;

/**
 * 导出基础处理,不涉及POI,只涉及对象,保证复用性
 *
 * @author JueYue 2014年8月9日 下午11:01:32
 */
@SuppressWarnings("rawtypes")
public class ExportCommonService {

    protected static final Logger LOGGER = LoggerFactory.getLogger(ExportCommonService.class);

    protected IExcelDataHandler dataHandler;
    protected IExcelDictHandler dictHandler;
    protected IExcelI18nHandler i18nHandler;

    protected List<String> needHandlerList;

    /**
     * 创建导出实体对象
     */
    private ExcelExportEntity createExcelExportEntity(Field field, String targetId,
                                                      Class<?> pojoClass,
                                                      List<Method> getMethods, ExcelEntity excelGroup) throws Exception {
        Excel excel = field.getAnnotation(Excel.class);
        ExcelExportEntity excelEntity = new ExcelExportEntity();
        excelEntity.setType(excel.type());
        getExcelField(targetId, field, excelEntity, excel, pojoClass, excelGroup);
        if (getMethods != null) {
            List<Method> newMethods = new ArrayList<Method>();
            newMethods.addAll(getMethods);
            newMethods.add(excelEntity.getMethod());
            excelEntity.setMethods(newMethods);
        }
        return excelEntity;
    }

    private Object dateFormatValue(Object value, ExcelExportEntity entity) throws Exception {
        Date temp = null;
        if (value instanceof String && StringUtils.isNoneEmpty(value.toString())) {
            SimpleDateFormat format = new SimpleDateFormat(entity.getDatabaseFormat());
            temp = format.parse(value.toString());
        } else if (value instanceof Date) {
            temp = (Date) value;
        } else if (value instanceof Instant) {
            Instant instant = (Instant)value;
            temp = Date.from(instant);
        }  else if (value instanceof LocalDate) {
            LocalDate localDate = (LocalDate)value;
            temp = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
        } else if(value instanceof LocalDateTime){
            LocalDateTime localDateTime = (LocalDateTime)value;
            temp = Date.from(localDateTime.atZone(ZoneId.systemDefault()).toInstant());
        } else if(value instanceof java.sql.Date) {
            temp = new Date(((java.sql.Date) value).getTime());
        } else if (value instanceof java.sql.Time) {
            temp = new Date(((java.sql.Time) value).getTime());
        } else if (value instanceof java.sql.Timestamp) {
            temp = new Date(((java.sql.Timestamp) value).getTime());
        }
        if (temp != null) {
            SimpleDateFormat format = new SimpleDateFormat(entity.getFormat());
            if(StringUtils.isNotEmpty(entity.getTimezone())){
                format.setTimeZone(TimeZone.getTimeZone(entity.getTimezone()));
            }
            value = format.format(temp);
        }
        return value;
    }


    private Object numFormatValue(Object value, ExcelExportEntity entity) {
        if (value == null) {
            return null;
        }
        if (!NumberUtils.isNumber(value.toString())) {
            LOGGER.error("data want num format ,but is not num, value is:" + value);
            return null;
        }
        Double d = Double.parseDouble(value.toString());
        DecimalFormat df = new DecimalFormat(entity.getNumFormat());
        return df.format(d);
    }

    /**
     * 获取需要导出的全部字段
     *
     * @param targetId 目标ID
     */
    public void getAllExcelField(String[] exclusions, String targetId, Field[] fields,
                                 List<ExcelExportEntity> excelParams, Class<?> pojoClass,
                                 List<Method> getMethods, ExcelEntity excelGroup) throws Exception {
        List<String> exclusionsList = exclusions != null ? Arrays.asList(exclusions) : null;
        ExcelExportEntity excelEntity;
        // 遍历整个filed
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            // 先判断是不是collection,在判断是不是java自带对象,之后就是我们自己的对象了
            if (PoiPublicUtil.isNotUserExcelUserThis(exclusionsList, field, targetId)) {
                continue;
            }
            // 首先判断Excel 可能一下特殊数据用户回自定义处理
            if (field.getAnnotation(Excel.class) != null) {
                Excel excel = field.getAnnotation(Excel.class);
                String name = PoiPublicUtil.getValueByTargetId(excel.name(), targetId, null);
                if (StringUtils.isNotBlank(name)) {
                    excelParams.add(createExcelExportEntity(field, targetId, pojoClass, getMethods, excelGroup));
                }
            } else if (PoiPublicUtil.isCollection(field.getType())) {
                ExcelCollection excel = field.getAnnotation(ExcelCollection.class);
                ParameterizedType pt = (ParameterizedType) field.getGenericType();
                Class<?> clz = (Class<?>) pt.getActualTypeArguments()[0];
                List<ExcelExportEntity> list = new ArrayList<ExcelExportEntity>();
                getAllExcelField(exclusions,
                        StringUtils.isNotEmpty(excel.id()) ? excel.id() : targetId,
                        PoiPublicUtil.getClassFields(clz), list, clz, null, null);
                excelEntity = new ExcelExportEntity();
                excelEntity.setName(PoiPublicUtil.getValueByTargetId(excel.name(), targetId, null));
                if (i18nHandler != null) {
                    excelEntity.setName(i18nHandler.getLocaleName(excelEntity.getName()));
                }
                excelEntity.setOrderNum(Integer
                        .valueOf(PoiPublicUtil.getValueByTargetId(excel.orderNum(), targetId, "0")));
                excelEntity
                        .setMethod(PoiReflectorUtil.fromCache(pojoClass).getGetMethod(field.getName()));
                excelEntity.setList(list);
                excelParams.add(excelEntity);
            } else {
                List<Method> newMethods = new ArrayList<Method>();
                if (getMethods != null) {
                    newMethods.addAll(getMethods);
                }
                newMethods.add(PoiReflectorUtil.fromCache(pojoClass).getGetMethod(field.getName()));
                ExcelEntity excel = field.getAnnotation(ExcelEntity.class);
                if (excel.show() && StringUtils.isEmpty(excel.name())) {
                    throw new ExcelExportException("if use ExcelEntity ,name mus has value ,data: " + ReflectionToStringBuilder.toString(excel), ExcelExportEnum.PARAMETER_ERROR);
                }
                getAllExcelField(exclusions,
                        StringUtils.isNotEmpty(excel.id()) ? excel.id() : targetId,
                        PoiPublicUtil.getClassFields(field.getType()), excelParams, field.getType(),
                        newMethods, excel.show() ? excel : null);
            }
        }
    }

    /**
     * 获取填如这个cell的值,提供一些附加功能
     */
    @SuppressWarnings("unchecked")
    public Object getCellValue(ExcelExportEntity entity, Object obj) throws Exception {
        Object value;
        if (obj instanceof Map) {
            value = ((Map<?, ?>) obj).get(entity.getKey());
        } else {
            value = entity.getMethods() != null ? getFieldBySomeMethod(entity.getMethods(), obj)
                    : entity.getMethod().invoke(obj, new Object[]{});
        }
        if (StringUtils.isNotEmpty(entity.getFormat())) {
            value = dateFormatValue(value, entity);
        }
        if (entity.getReplace() != null && entity.getReplace().length > 0) {
            value = replaceValue(entity.getReplace(), String.valueOf(value));
        }
        if (StringUtils.isNotEmpty(entity.getNumFormat())) {
            value = numFormatValue(value, entity);
        }
        if (StringUtils.isNotEmpty(entity.getDict()) && dictHandler != null) {
            value = dictHandler.toName(entity.getDict(), obj, entity.getName(), value);
        }
        if (needHandlerList != null && needHandlerList.contains(entity.getName())) {
            value = dataHandler.exportHandler(obj, entity.getName(), value);
        }
        if (StringUtils.isNotEmpty(entity.getSuffix()) && value != null) {
            value = value + entity.getSuffix();
        }
        if (value != null && StringUtils.isNotEmpty(entity.getEnumExportField())) {
            value = PoiReflectorUtil.fromCache(value.getClass()).getValue(value, entity.getEnumExportField());
        }
        return value == null ? "" : value.toString();
    }


    /**
     * 获取集合的值
     */
    public Collection<?> getListCellValue(ExcelExportEntity entity, Object obj) throws Exception {
        Object value;
        if (obj instanceof Map) {
            value = ((Map<?, ?>) obj).get(entity.getKey());
        } else {
            value = (Collection<?>) entity.getMethod().invoke(obj, new Object[]{});
        }
        return (Collection<?>) value;
    }

    /**
     * 注解到导出对象的转换
     */
    private void getExcelField(String targetId, Field field, ExcelExportEntity excelEntity,
                               Excel excel, Class<?> pojoClass, ExcelEntity excelGroup) throws Exception {
        excelEntity.setName(PoiPublicUtil.getValueByTargetId(excel.name(), targetId, null));
        excelEntity.setWidth(excel.width());
        excelEntity.setHeight(excel.height());
        excelEntity.setNeedMerge(excel.needMerge());
        excelEntity.setMergeVertical(excel.mergeVertical());
        excelEntity.setMergeRely(excel.mergeRely());
        excelEntity.setReplace(excel.replace());
        excelEntity.setOrderNum(
                Integer.valueOf(PoiPublicUtil.getValueByTargetId(excel.orderNum(), targetId, "0")));
        excelEntity.setWrap(excel.isWrap());
        excelEntity.setExportImageType(excel.imageType());
        excelEntity.setSuffix(excel.suffix());
        excelEntity.setDatabaseFormat(excel.databaseFormat());
        excelEntity.setFormat(
                StringUtils.isNotEmpty(excel.exportFormat()) ? excel.exportFormat() : excel.format());
        excelEntity.setStatistics(excel.isStatistics());
        excelEntity.setHyperlink(excel.isHyperlink());
        excelEntity.setMethod(PoiReflectorUtil.fromCache(pojoClass).getGetMethod(field.getName()));
        excelEntity.setNumFormat(excel.numFormat());
        excelEntity.setColumnHidden(excel.isColumnHidden());
        excelEntity.setDict(excel.dict());
        excelEntity.setEnumExportField(excel.enumExportField());
        excelEntity.setTimezone(excel.timezone());
        if (excelGroup != null) {
            excelEntity.setGroupName(PoiPublicUtil.getValueByTargetId(excelGroup.name(), targetId, null));
        } else {
            excelEntity.setGroupName(excel.groupName());
        }
        if (i18nHandler != null) {
            excelEntity.setName(i18nHandler.getLocaleName(excelEntity.getName()));
            excelEntity.setGroupName(i18nHandler.getLocaleName(excelEntity.getGroupName()));
        }
    }

    /**
     * 多个反射获取值
     */
    public Object getFieldBySomeMethod(List<Method> list, Object t) throws Exception {
        for (Method m : list) {
            if (t == null) {
                t = "";
                break;
            }
            t = m.invoke(t, new Object[]{});
        }
        return t;
    }

    /**
     * 根据注解获取行高
     */
    public short getRowHeight(List<ExcelExportEntity> excelParams) {
        double maxHeight = 0;
        for (int i = 0; i < excelParams.size(); i++) {
            maxHeight = maxHeight > excelParams.get(i).getHeight() ? maxHeight
                    : excelParams.get(i).getHeight();
            if (excelParams.get(i).getList() != null) {
                for (int j = 0; j < excelParams.get(i).getList().size(); j++) {
                    maxHeight = maxHeight > excelParams.get(i).getList().get(j).getHeight()
                            ? maxHeight : excelParams.get(i).getList().get(j).getHeight();
                }
            }
        }
        return (short) (maxHeight * 50);
    }

    private Object replaceValue(String[] replace, String value) {
        String[] temp;
        for (String str : replace) {
            temp = str.split("_");
            if (value.equals(temp[1])) {
                value = temp[0];
                break;
            }
        }
        return value;
    }

    /**
     * 对字段根据用户设置排序
     */
    public void sortAllParams(List<ExcelExportEntity> excelParams) {
        // 自然排序,group 内部排序,集合内部排序
        // 把有groupName的统一收集起来,内部先排序
        Map<String, List<ExcelExportEntity>> groupMap = new HashMap<String, List<ExcelExportEntity>>();
        for (int i = excelParams.size() - 1; i > -1; i--) {
            // 集合内部排序
            if (excelParams.get(i).getList() != null) {
                Collections.sort(excelParams.get(i).getList());
            } else if (StringUtils.isNoneEmpty(excelParams.get(i).getGroupName())) {
                if (!groupMap.containsKey(excelParams.get(i).getGroupName())) {
                    groupMap.put(excelParams.get(i).getGroupName(), new ArrayList<ExcelExportEntity>());
                }
                groupMap.get(excelParams.get(i).getGroupName()).add(excelParams.get(i));
                excelParams.remove(i);
            }
        }
        Collections.sort(excelParams);
        if (groupMap.size() > 0) {
            // group 内部排序
            for (Iterator it = groupMap.entrySet().iterator(); it.hasNext(); ) {
                Map.Entry<String, List<ExcelExportEntity>> entry = (Map.Entry) it.next();
                Collections.sort(entry.getValue());
                // 插入到excelParams当中
                boolean isInsert = false;
                String groupName = "START";
                for (int i = 0; i < excelParams.size(); i++) {
                    // 跳过groupName 的元素,防止破会内部结构
                    if (excelParams.get(i).getOrderNum() > entry.getValue().get(0).getOrderNum()
                            && !groupName.equals(excelParams.get(i).getGroupName())) {
                        if (StringUtils.isNotEmpty(excelParams.get(i).getGroupName())) {
                            groupName = excelParams.get(i).getGroupName();
                        }
                        excelParams.addAll(i, entry.getValue());
                        isInsert = true;
                        break;
                    } else if (!groupName.equals(excelParams.get(i).getGroupName()) &&
                            StringUtils.isNotEmpty(excelParams.get(i).getGroupName())) {
                        groupName = excelParams.get(i).getGroupName();
                    }
                }
                //如果都比他小就插入到最后
                if (!isInsert) {
                    excelParams.addAll(entry.getValue());
                }
            }
        }
    }

    /**
     * 添加Index列
     */
    public ExcelExportEntity indexExcelEntity(ExportParams entity) {
        ExcelExportEntity exportEntity = new ExcelExportEntity();
        //保证是第一排
        exportEntity.setOrderNum(Integer.MIN_VALUE);
        exportEntity.setNeedMerge(true);
        exportEntity.setName(entity.getIndexName());
        exportEntity.setWidth(10);
        exportEntity.setFormat(PoiBaseConstants.IS_ADD_INDEX);
        return exportEntity;
    }

    /**
     * 获取导出报表的字段总长度
     */
    public int getFieldLength(List<ExcelExportEntity> excelParams) {
        int length = -1;// 从0开始计算单元格的
        for (ExcelExportEntity entity : excelParams) {
            if (entity.getList() != null) {
                length += getFieldLength(entity.getList()) + 1;
            } else {
                length++;
            }
        }
        return length;
    }

    /**
     * 判断表头是只有一行还是多行
     */
    public int getRowNums(List<ExcelExportEntity> excelParams,boolean isDeep) {
        for (int i = 0; i < excelParams.size(); i++) {
            if (excelParams.get(i).getList() != null
                    && StringUtils.isNotBlank(excelParams.get(i).getName())) {
                return isDeep? 1 + getRowNums(excelParams.get(i).getList() , isDeep) : 2;
            }
            if (StringUtils.isNotEmpty(excelParams.get(i).getGroupName())) {
                return 2;
            }
        }
        return 1;
    }

}
