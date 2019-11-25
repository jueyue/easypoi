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
package cn.afterturn.easypoi.excel;

import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.export.ExcelBatchExportService;
import cn.afterturn.easypoi.excel.export.ExcelExportService;
import cn.afterturn.easypoi.excel.export.template.ExcelExportOfTemplateUtil;
import cn.afterturn.easypoi.handler.inter.IExcelExportServer;
import cn.afterturn.easypoi.handler.inter.IWriter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * excel 导出工具类
 *
 * @author JueYue
 * @version 1.0
 * 2013-10-17
 */
public final class ExcelExportUtil {

    public static       int    USE_SXSSF_LIMIT = 100000;
    public static final String SHEET_NAME      = "sheetName";

    private ExcelExportUtil() {
    }

    /**
     * 大数据量导出
     *
     * @param entity    表格标题属性
     * @param pojoClass Excel对象Class
     */
    public static IWriter<Workbook> exportBigExcel(ExportParams entity, Class<?> pojoClass) {
        ExcelBatchExportService batchServer = new ExcelBatchExportService();
        batchServer.init(entity, pojoClass);
        return batchServer;
    }

    /**
     * 大数据量导出
     *
     * @param entity
     * @param excelParams
     * @return
     */
    public static IWriter<Workbook> exportBigExcel(ExportParams entity, List<ExcelExportEntity> excelParams) {
        ExcelBatchExportService batchServer = new ExcelBatchExportService();
        batchServer.init(entity, excelParams);
        return batchServer;
    }

    /**
     * 大数据量导出
     *
     * @param entity      表格标题属性
     * @param pojoClass   Excel对象Class
     * @param server      查询数据的接口
     * @param queryParams 查询数据的参数
     */
    public static Workbook exportBigExcel(ExportParams entity, Class<?> pojoClass,
                                          IExcelExportServer server, Object queryParams) {
        ExcelBatchExportService batchServer = new ExcelBatchExportService();
        batchServer.init(entity, pojoClass);
        return batchServer.exportBigExcel(server, queryParams);
    }

    /**
     * 大数据量导出
     *
     * @param entity
     * @param excelParams
     * @param server      查询数据的接口
     * @param queryParams 查询数据的参数
     * @return
     */
    public static Workbook exportBigExcel(ExportParams entity, List<ExcelExportEntity> excelParams,
                                          IExcelExportServer server, Object queryParams) {
        ExcelBatchExportService batchServer = new ExcelBatchExportService();
        batchServer.init(entity, excelParams);
        return batchServer.exportBigExcel(server, queryParams);
    }


    /**
     * @param entity    表格标题属性
     * @param pojoClass Excel对象Class
     * @param dataSet   Excel对象数据List
     */
    public static Workbook exportExcel(ExportParams entity, Class<?> pojoClass,
                                       Collection<?> dataSet) {
        Workbook workbook = getWorkbook(entity.getType(), dataSet.size());
        new ExcelExportService().createSheet(workbook, entity, pojoClass, dataSet);
        return workbook;
    }

    private static Workbook getWorkbook(ExcelType type, int size) {
        if (ExcelType.HSSF.equals(type)) {
            return new HSSFWorkbook();
        } else if (size < USE_SXSSF_LIMIT) {
            return new XSSFWorkbook();
        } else {
            return new SXSSFWorkbook();
        }
    }

    /**
     * 根据Map创建对应的Excel
     *
     * @param entity     表格标题属性
     * @param entityList Map对象列表
     * @param dataSet    Excel对象数据List
     */
    public static Workbook exportExcel(ExportParams entity, List<ExcelExportEntity> entityList,
                                       Collection<?> dataSet) {
        Workbook workbook = getWorkbook(entity.getType(), dataSet.size());
        ;
        new ExcelExportService().createSheetForMap(workbook, entity, entityList, dataSet);
        return workbook;
    }

    /**
     * 根据Map创建对应的Excel(一个excel 创建多个sheet)
     *
     * @param list 多个Map key title 对应表格Title key entity 对应表格对应实体 key data
     *             Collection 数据
     * @return
     */
    public static Workbook exportExcel(List<Map<String, Object>> list, ExcelType type) {
        Workbook workbook = getWorkbook(type, 0);
        for (Map<String, Object> map : list) {
            ExcelExportService service = new ExcelExportService();
            service.createSheet(workbook, (ExportParams) map.get("title"),
                    (Class<?>) map.get("entity"), (Collection<?>) map.get("data"));
        }
        return workbook;
    }

    /**
     * 导出文件通过模板解析,不推荐这个了,推荐全部通过模板来执行处理
     *
     * @param params    导出参数类
     * @param pojoClass 对应实体
     * @param dataSet   实体集合
     * @param map       模板集合
     * @return
     */
    @Deprecated
    public static Workbook exportExcel(TemplateExportParams params, Class<?> pojoClass,
                                       Collection<?> dataSet, Map<String, Object> map) {
        return new ExcelExportOfTemplateUtil().createExcelByTemplate(params, pojoClass, dataSet,
                map);
    }

    /**
     * 导出文件通过模板解析只有模板,没有集合
     *
     * @param params 导出参数类
     * @param map    模板集合
     * @return
     */
    public static Workbook exportExcel(TemplateExportParams params, Map<String, Object> map) {
        return new ExcelExportOfTemplateUtil().createExcelByTemplate(params, null, null, map);
    }

    /**
     * 导出文件通过模板解析只有模板,没有集合
     * 每个sheet对应一个map,导出到处,key是sheet的NUM
     *
     * @param params 导出参数类
     * @param map    模板集合
     * @return
     */
    public static Workbook exportExcel(Map<Integer, Map<String, Object>> map,
                                       TemplateExportParams params) {
        return new ExcelExportOfTemplateUtil().createExcelByTemplate(params, map);
    }

    /**
     * 导出文件通过模板解析只有模板,没有集合
     * 每个sheet对应一个list,按照数量进行导出排序,key是sheet的NUM
     *
     * @param params 导出参数类
     * @param map    模板集合
     * @return
     */
    public static Workbook exportExcelClone(Map<Integer, List<Map<String, Object>>> map,
                                            TemplateExportParams params) {
        return new ExcelExportOfTemplateUtil().createExcelCloneByTemplate(params, map);
    }

}
