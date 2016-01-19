package org.jeecgframework.poi.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.cache.HtmlCache;
import org.jeecgframework.poi.excel.entity.ExcelToHtmlParams;

/**
 * Excel 变成界面
 * @author JueYue
 * @date 2015年5月10日 上午11:51:48
 */
public class ExcelToHtmlUtil {

    private ExcelToHtmlUtil() {
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
     * 转换成为Table,显示图片
     * @param wb Excel
     * @return
     */
    public static String toTableHtml(Workbook wb, String path) {
        return HtmlCache.getHtml(new ExcelToHtmlParams(wb, false, 0, path));
    }

    /**
     * 转换成为Table
     * @param wb Excel
     * @param sheetNum sheetNum
     * @return
     */
    public static String toTableHtml(Workbook wb, int sheetNum) {
        return HtmlCache.getHtml(new ExcelToHtmlParams(wb, false, sheetNum, null));
    }

    /**
     * 转换成为Table,显示图片
     * @param wb Excel
     * @param sheetNum sheetNum
     * @return
     */
    public static String toTableHtml(Workbook wb, int sheetNum, String path) {
        return HtmlCache.getHtml(new ExcelToHtmlParams(wb, false, sheetNum, path));
    }

    /**
     * 转换成为完整界面
     * @param wb Excel
     * @return
     */
    public static String toAllHtml(Workbook wb) {
        return HtmlCache.getHtml(new ExcelToHtmlParams(wb, true, 0, null));
    }

    /**
     * 转换成为完整界面,显示图片
     * @param wb Excel
     * @param path 图片保存路径
     * @return
     */
    public static String toAllHtml(Workbook wb, String path) {
        return HtmlCache.getHtml(new ExcelToHtmlParams(wb, true, 0, path));
    }

    /**
     * 转换成为完整界面
     * @param wb Excel
     * @param sheetNum sheetNum
     * @return
     */
    public static String toAllHtml(Workbook wb, int sheetNum) {
        return HtmlCache.getHtml(new ExcelToHtmlParams(wb, true, sheetNum, null));
    }

    /**
     * 转换成为完整界面,显示图片
     * @param wb Excel
     * @param sheetNum sheetNum
     * @return
     */
    public static String toAllHtml(Workbook wb, int sheetNum, String path) {
        return HtmlCache.getHtml(new ExcelToHtmlParams(wb, true, sheetNum, path));
    }

}
