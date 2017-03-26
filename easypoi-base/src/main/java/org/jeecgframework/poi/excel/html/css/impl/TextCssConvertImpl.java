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
import org.jeecgframework.poi.excel.html.css.ICssConvertToExcel;
import org.jeecgframework.poi.excel.html.css.ICssConvertToHtml;
import org.jeecgframework.poi.excel.html.entity.style.CellStyleEntity;

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
        int color = style.getFont().getColor();
        if (color != 0 && color != BLACK.index) {
            font.setColor((short) color);
        }
        if (UNDERLINE.equals(style.getFont().getDecoration())) {
            font.setUnderline(Font.U_SINGLE);
        }
        cellStyle.setFont(font);
    }

}
