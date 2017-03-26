package org.jeecgframework.poi.excel.html.css.impl;

import static org.jeecgframework.poi.excel.html.entity.HtmlCssConstant.*;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.BLACK;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.jeecgframework.poi.excel.html.css.ICssConvertToExcel;
import org.jeecgframework.poi.excel.html.css.ICssConvertToHtml;
import org.jeecgframework.poi.excel.html.entity.style.CellStyleEntity;
import org.jeecgframework.poi.util.PoiCssUtils;

public class TextCssConvertImpl implements ICssConvertToExcel, ICssConvertToHtml {

    @Override
    public String convertToHtml(Cell cell, CellStyle cellStyle, CellStyleEntity style) {

        return null;
    }

    @Override
    public void convertToExcel(Cell cell, CellStyle cellStyle, CellStyleEntity style) {
        if (style == null || style.getFont() == null) {
            return;
        }
        Font font = cell.getSheet().getWorkbook().createFont();
        if (ITALIC.equals(style.getFont().getStyle())) {
            font.setItalic(true);
        }
        int fontSize = style.getFont().getSize();
        if (fontSize > 0) {
            font.setFontHeightInPoints((short) fontSize);
        }
        if (BOLD.equals(style.getFont().getWeight())) {
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        }
        String fontFamily = style.getFont().getFamily();
        if (StringUtils.isNotBlank(fontFamily)) {
            font.setFontName(fontFamily);
        }
        String color = style.getFont().getColor();
        if (StringUtils.isNoneEmpty(color)) {
            if (font instanceof HSSFFont) {
                setFoutForHSSF(font, cell.getSheet().getWorkbook(), color);
            } else if (font instanceof XSSFFont) {
                setFoutForXSSF(font, color);
            }
        }
        if (UNDERLINE.equals(style.getFont().getDecoration())) {
            font.setUnderline(Font.U_SINGLE);
        }
        cellStyle.setFont(font);
    }

    private void setFoutForXSSF(Font font, String colorStr) {
        XSSFColor color = PoiCssUtils.parseColor(colorStr);
        ((XSSFFont)font).setColor(color);
    }

    private void setFoutForHSSF(Font font, Workbook workbook, String colorStr) {
        HSSFColor color = PoiCssUtils.parseColor((HSSFWorkbook) workbook, colorStr);
        if (color != null) {
            if (color.getIndex() != BLACK.index) {
                font.setColor(color.getIndex());
            }
        }
    }

}
