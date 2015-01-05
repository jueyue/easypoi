package org.jeecgframework.poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.jeecgframework.poi.excel.entity.ImportParams;
import org.jeecgframework.poi.excel.entity.result.ExcelImportResult;
import org.jeecgframework.poi.excel.imports.ExcelImportServer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Excel 导入工具
 * 
 * @author JueYue
 * @date 2013-9-24
 * @version 1.0
 */
@SuppressWarnings({ "unchecked" })
public class ExcelImportUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelImportUtil.class);

    /**
     * Excel 导入 数据源本地文件,不返回校验结果 导入 字 段类型 Integer,Long,Double,Date,String,Boolean
     * 
     * @param file
     * @param pojoClass
     * @param params
     * @return
     * @throws Exception
     */
    public static <T> List<T> importExcel(File file, Class<?> pojoClass, ImportParams params) {
        FileInputStream in = null;
        List<T> result = null;
        try {
            in = new FileInputStream(file);
            result = new ExcelImportServer().importExcelByIs(in, pojoClass, params).getList();
        } catch (Exception e) {
            LOGGER.error(e.getMessage(),e);
        } finally {
            try {
                in.close();
            } catch (IOException e) {
                LOGGER.error(e.getMessage(),e);
            }
        }
        return result;
    }

    /**
     * Excel 导入 数据源IO流,不返回校验结果 导入 字段类型 Integer,Long,Double,Date,String,Boolean
     * 
     * @param file
     * @param pojoClass
     * @param params
     * @return
     * @throws Exception
     */
    public static <T> List<T> importExcelByIs(InputStream inputstream, Class<?> pojoClass,
                                              ImportParams params) throws Exception {
        return new ExcelImportServer().importExcelByIs(inputstream, pojoClass, params).getList();
    }

    /**
     * Excel 导入 数据源IO流,返回校验结果 字段类型 Integer,Long,Double,Date,String,Boolean
     * 
     * @param file
     * @param pojoClass
     * @param params
     * @return
     * @throws Exception
     */
    public static ExcelImportResult importExcelByIsAndVerify(InputStream inputstream,
                                                             Class<?> pojoClass, ImportParams params)
                                                                                                     throws Exception {
        return new ExcelImportServer().importExcelByIs(inputstream, pojoClass, params);
    }

    /**
     * Excel 导入 数据源本地文件,返回校验结果 字段类型 Integer,Long,Double,Date,String,Boolean
     * 
     * @param file
     * @param pojoClass
     * @param params
     * @return
     * @throws Exception
     */
    public static ExcelImportResult importExcelVerify(File file, Class<?> pojoClass,
                                                      ImportParams params) {
        FileInputStream in = null;
        try {
            in = new FileInputStream(file);
            return new ExcelImportServer().importExcelByIs(in, pojoClass, params);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(),e);
        } finally {
            try {
                in.close();
            } catch (IOException e) {
                LOGGER.error(e.getMessage(),e);
            }
        }
        return null;
    }

}
