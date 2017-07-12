package cn.afterturn.easypoi.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;

import cn.afterturn.easypoi.cache.HtmlCache;
import cn.afterturn.easypoi.excel.entity.ExcelToHtmlParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.html.HtmlToExcelServer;
import cn.afterturn.easypoi.exception.excel.ExcelExportException;
import cn.afterturn.easypoi.exception.excel.enums.ExcelExportEnum;

/**
 * Excel 变成界面
 * @author JueYue
 *  2015年5月10日 上午11:51:48
 */
public class ExcelXorHtmlUtil {

    private ExcelXorHtmlUtil() {
    }

    /**
     * 转换成为Table
     * @param wb Excel
     * @return
     */
    public static String toTableHtml(Workbook wb) {
        return HtmlCache.getHtml(new ExcelToHtmlParams(wb, false, 0, null));
    }

    /**
     * 转换成为完整界面,显示图片
     * @param wb Excel
     * @return
     */
    public static String toAllHtml(Workbook wb) {
        return HtmlCache.getHtml(new ExcelToHtmlParams(wb, true, 0, null));
    }

    /**
     * 转换成为Table
     * @param params
     * @return
     */
    public static String excelToHtml(ExcelToHtmlParams params) {
        return HtmlCache.getHtml(params);
    }

    /**
     * Html 读取Excel
     * @param html
     * @param type
     * @return
     */
    public static Workbook htmlToExcel(String html, ExcelType type) {
        Workbook workbook = null;
        if (ExcelType.HSSF.equals(type)) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }
        new HtmlToExcelServer().createSheet(html, workbook);
        return workbook;
    }

    /**
     * Html 读取Excel
     * @param is
     * @param type
     * @return
     */
    public static Workbook htmlToExcel(InputStream is, ExcelType type) {
        try {
            return htmlToExcel(new String(IOUtils.toByteArray(is)), type);
        } catch (IOException e) {
            throw new ExcelExportException(ExcelExportEnum.HTML_ERROR, e);
        }
    }

}
