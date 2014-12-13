package org.jeecgframework.poi.excel;

import java.util.Collection;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.entity.vo.PoiBaseConstants;
import org.jeecgframework.poi.excel.export.ExcelExportServer;
import org.jeecgframework.poi.excel.export.template.ExcelExportOfTemplateUtil;

/**
 * excel 导出工具类
 * 
 * @author jueyue
 * @version 1.0
 * @date 2013-10-17
 */
public final class ExcelExportUtil {

    /**
     * @param entity
     *            表格标题属性
     * @param pojoClass
     *            Excel对象Class
     * @param dataSet
     *            Excel对象数据List
     */
    public static Workbook exportExcel(ExportParams entity, Class<?> pojoClass,
                                       Collection<?> dataSet) {
        Workbook workbook;
        if (entity.getType().equals(PoiBaseConstants.HSSF)) {
            workbook = new HSSFWorkbook();
        } else if (dataSet.size() < 1000) {
            workbook = new XSSFWorkbook();
        } else {
            workbook = new SXSSFWorkbook();
        }
        new ExcelExportServer().createSheet(workbook, entity, pojoClass, dataSet, entity.getType());
        return workbook;
    }

    /**
     * 根据Map创建对应的Excel
     * @param entity
     *            表格标题属性
     * @param pojoClass
     *            Excel对象Class
     * @param dataSet
     *            Excel对象数据List
     */
    public static Workbook exportExcel(ExportParams entity, List<ExcelExportEntity> entityList,
                                       Collection<? extends Map<?, ?>> dataSet) {
        Workbook workbook;
        if (entity.getType().equals(PoiBaseConstants.HSSF)) {
            workbook = new HSSFWorkbook();
        } else if (dataSet.size() < 1000) {
            workbook = new XSSFWorkbook();
        } else {
            workbook = new SXSSFWorkbook();
        }
        new ExcelExportServer().createSheetForMap(workbook, entity, entityList, dataSet,
            entity.getType());
        return workbook;
    }

    /**
     * 一个excel 创建多个sheet
     * 
     * @param list
     *            多个Map key title 对应表格Title key entity 对应表格对应实体 key data
     *            Collection 数据
     * @return
     */
    public static Workbook exportExcel(List<Map<String, Object>> list, String type) {
        Workbook workbook;
        if (type.equals(PoiBaseConstants.HSSF)) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }
        ExcelExportServer server = new ExcelExportServer();
        for (Map<String, Object> map : list) {
            server.createSheet(workbook, (ExportParams) map.get("title"),
                (Class<?>) map.get("entity"), (Collection<?>) map.get("data"), type);
        }
        return workbook;
    }

    /**
     * 导出文件通过模板解析
     * 
     * @param params
     *            导出参数类
     * @param pojoClass
     *            对应实体
     * @param dataSet
     *            实体集合
     * @param map
     *            模板集合
     * @return
     */
    public static Workbook exportExcel(TemplateExportParams params, Class<?> pojoClass,
                                       Collection<?> dataSet, Map<String, Object> map) {
        return new ExcelExportOfTemplateUtil().createExcleByTemplate(params, pojoClass, dataSet,
            map);
    }

    /**
     * 导出文件通过模板解析只有模板,没有集合
     * 
     * @param params
     *            导出参数类
     * @param map
     *            模板集合
     * @return
     */
    public static Workbook exportExcel(TemplateExportParams params, Map<String, Object> map) {
        return new ExcelExportOfTemplateUtil().createExcleByTemplate(params, null, null, map);
    }

}
