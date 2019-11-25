package cn.afterturn.easypoi.csv;

import cn.afterturn.easypoi.csv.entity.CsvExportParams;
import cn.afterturn.easypoi.csv.export.CsvExportService;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.handler.inter.IExcelExportServer;
import cn.afterturn.easypoi.handler.inter.IWriter;

import java.io.OutputStream;
import java.util.Collection;
import java.util.List;

/**
 * Csv批量导出文件
 *
 * @author by jueyue on 18-11-14.
 */
public final class CsvExportUtil {

    private CsvExportUtil() {
    }


    /**
     * export csv all list
     *
     * @param params
     * @param pojoClass
     * @param dataSet
     * @param outputStream
     */
    public static void exportCsv(CsvExportParams params, Class<?> pojoClass, Collection<?> dataSet, OutputStream outputStream) {
        IWriter writer = new CsvExportService(outputStream, params, pojoClass);
        writer.writer(dataSet);
        writer.close();
    }

    /**
     * export csv use server
     *
     * @param params
     * @param pojoClass
     * @param server
     * @param queryParams
     * @param outputStream
     */
    public static void exportCsv(CsvExportParams params, Class<?> pojoClass, IExcelExportServer server, Object queryParams, OutputStream outputStream) {
        IWriter    writer = new CsvExportService(outputStream, params, pojoClass);
        int        page   = 1;
        Collection dataSet;
        while ((dataSet = server.selectListForExcelExport(queryParams, page)) != null && dataSet.size() > 0) {
            page++;
            writer.writer(dataSet);
        }
        writer.close();
    }

    /**
     * @param params    表格标题属性
     * @param pojoClass Excel对象Class
     */
    public static IWriter<Void> exportCsv(CsvExportParams params, Class<?> pojoClass, OutputStream outputStream) {
        CsvExportUtil ce = new CsvExportUtil();
        return new CsvExportService(outputStream, params, pojoClass);
    }

    /**
     * 根据Map创建对应的Excel
     *
     * @param params     表格标题属性
     * @param entityList Map对象列表
     */
    public static IWriter<Void> exportCsv(CsvExportParams params, List<ExcelExportEntity> entityList, OutputStream outputStream) {
        return new CsvExportService(outputStream, params, entityList);
    }

}
