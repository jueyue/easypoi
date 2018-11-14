package cn.afterturn.easypoi.csv;

import cn.afterturn.easypoi.csv.entity.CsvExportParams;
import cn.afterturn.easypoi.csv.export.CsvExportService;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.export.ExcelBatchExportService;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.OutputStream;
import java.util.Collection;
import java.util.List;

/**
 * Csv批量导出文件
 *
 * @author by jueyue on 18-11-14.
 */
public final class CsvExportUtil {

    /**
     * @param entity    表格标题属性
     * @param pojoClass Excel对象Class
     * @param dataSet   Excel对象数据List
     */
    public static Workbook exportBigExcel(ExportParams entity, Class<?> pojoClass,
                                          Collection<?> dataSet) {
        ExcelBatchExportService batchService = ExcelBatchExportService
                .getExcelBatchExportService(entity, pojoClass);
        return batchService.appendData(dataSet);
    }

    public static Workbook exportBigExcel(ExportParams entity, List<ExcelExportEntity> excelParams,
                                          Collection<?> dataSet) {
        ExcelBatchExportService batchService = ExcelBatchExportService
                .getExcelBatchExportService(entity, excelParams);
        return batchService.appendData(dataSet);
    }

    public static void closeExportBigExcel() {
        ExcelBatchExportService batchService = ExcelBatchExportService.getCurrentExcelBatchExportService();
        if (batchService != null) {
            batchService.closeExportBigExcel();
        }
    }

    /**
     * @param params    表格标题属性
     * @param pojoClass Excel对象Class
     * @param dataSet   Excel对象数据List
     */
    public static void exportCsv(CsvExportParams params, Class<?> pojoClass,
                                 Collection<?> dataSet, OutputStream outputStream) {
        new CsvExportService().createCsv(outputStream, params, pojoClass, dataSet);
    }

    /**
     * 根据Map创建对应的Excel
     *
     * @param params     表格标题属性
     * @param entityList Map对象列表
     * @param dataSet    Excel对象数据List
     */
    public static void exportCsv(CsvExportParams params, List<ExcelExportEntity> entityList,
                                 Collection<?> dataSet, OutputStream outputStream) {
        new CsvExportService().createCsvOfList(outputStream, params, entityList, dataSet);
    }
}
