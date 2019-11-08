package cn.afterturn.easypoi.csv;

import cn.afterturn.easypoi.csv.entity.CsvExportParams;
import cn.afterturn.easypoi.csv.export.CsvExportService;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.util.PoiZipUtil;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Collection;
import java.util.List;

/**
 * Csv批量导出文件
 *
 * @author by jueyue on 18-11-14.
 */
public final class CsvExportUtil {

    private CsvExportService cs;

    /**
     * @param params    表格标题属性
     * @param pojoClass Excel对象Class
     */
    public static CsvExportUtil exportCsv(CsvExportParams params, Class<?> pojoClass, OutputStream outputStream) {
        CsvExportUtil ce = new CsvExportUtil();
        ce.cs = new CsvExportService(outputStream, params, pojoClass);
        return ce;
    }

    /**
     * 根据Map创建对应的Excel
     *
     * @param params     表格标题属性
     * @param entityList Map对象列表
     */
    public static CsvExportUtil exportCsv(CsvExportParams params, List<ExcelExportEntity> entityList, OutputStream outputStream) {
        CsvExportUtil ce = new CsvExportUtil();
        ce.cs = new CsvExportService(outputStream, params, entityList);
        return ce;
    }

    public CsvExportUtil write(Collection<?> dataSet) {
        this.cs.write(dataSet);
        return this;
    }

    public void close() {
        this.cs.close();
    }
}
