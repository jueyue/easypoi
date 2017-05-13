package cn.afterturn.easypoi.excel.html.css.impl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

import cn.afterturn.easypoi.excel.html.css.ICssConvertToExcel;
import cn.afterturn.easypoi.excel.html.css.ICssConvertToHtml;
import cn.afterturn.easypoi.excel.html.entity.HtmlCssConstant;
import cn.afterturn.easypoi.excel.html.entity.style.CellStyleEntity;

public class AlignCssConvertImpl implements ICssConvertToExcel, ICssConvertToHtml {

    @Override
    public String convertToHtml(Cell cell, CellStyle cellStyle, CellStyleEntity style) {

        return null;
    }

    @Override
    public void convertToExcel(Cell cell, CellStyle cellStyle, CellStyleEntity style) {
        // align
        if (HtmlCssConstant.RIGHT.equals(style.getAlign())) {
            cellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
        } else if (HtmlCssConstant.CENTER.equals(style.getAlign())) {
            cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        } else if (HtmlCssConstant.LEFT.equals(style.getAlign())) {
            cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
        } else if (HtmlCssConstant.JUSTIFY.equals(style.getAlign())) {
            cellStyle.setAlignment(CellStyle.ALIGN_JUSTIFY);
        }
        // vertical align
        if (HtmlCssConstant.TOP.equals(style.getVetical())) {
            cellStyle.setVerticalAlignment(CellStyle.VERTICAL_TOP);
        } else if (HtmlCssConstant.CENTER.equals(style.getAlign())) {
            cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        } else if (HtmlCssConstant.BOTTOM.equals(style.getAlign())) {
            cellStyle.setVerticalAlignment(CellStyle.VERTICAL_BOTTOM);
        } else if (HtmlCssConstant.JUSTIFY.equals(style.getAlign())) {
            cellStyle.setVerticalAlignment(CellStyle.VERTICAL_JUSTIFY);
        }
    }

}
