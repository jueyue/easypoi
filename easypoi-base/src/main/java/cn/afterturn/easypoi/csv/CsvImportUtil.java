package cn.afterturn.easypoi.csv;

import cn.afterturn.easypoi.csv.entity.CsvImportParams;
import cn.afterturn.easypoi.csv.handler.ICsvSaveDataHandler;
import cn.afterturn.easypoi.csv.imports.CsvImportService;

import java.io.InputStream;
import java.util.List;

/**
 * CSV 导入工具类
 * 具体和Excel类似,但是比Excel简单
 * 需要处理一些字符串的处理
 *
 * @author by jueyue on 18-10-3.
 */
public final class CsvImportUtil {

    /**
     * Csv 导入流适合大数据导入
     * 导入 数据源IO流,不返回校验结果 导入 字段类型 Integer,Long,Double,Date,String,Boolean
     *
     * @param inputstream
     * @param pojoClass
     * @param params
     * @return
     */
    public static <T> List<T> importCsv(InputStream inputstream, Class<?> pojoClass,
                                        CsvImportParams params) {
        return new CsvImportService().readExcel(inputstream, pojoClass, params, null);
    }

    /**
     * Csv 导入流适合大数据导入
     * 导入 数据源IO流,不返回校验结果 导入 字段类型 Integer,Long,Double,Date,String,Boolean
     *
     * @param inputstream
     * @param pojoClass
     * @param params
     * @return
     */
    public static <T> List<T> importCsv(InputStream inputstream, Class<?> pojoClass,
                                        CsvImportParams params, ICsvSaveDataHandler saveDataHandler) {
        return new CsvImportService().readExcel(inputstream, pojoClass, params, saveDataHandler);
    }
}
