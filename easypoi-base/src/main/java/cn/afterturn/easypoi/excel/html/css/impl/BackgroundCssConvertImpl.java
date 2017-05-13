package cn.afterturn.easypoi.excel.html.css.impl;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import cn.afterturn.easypoi.excel.html.css.ICssConvertToExcel;
import cn.afterturn.easypoi.excel.html.css.ICssConvertToHtml;
import cn.afterturn.easypoi.excel.html.entity.style.CellStyleEntity;
import cn.afterturn.easypoi.util.PoiCssUtils;

public class BackgroundCssConvertImpl implements ICssConvertToExcel, ICssConvertToHtml {

    @Override
    public String convertToHtml(Cell cell, CellStyle cellStyle, CellStyleEntity style) {

        return null;
    }

    @Override
    public void convertToExcel(Cell cell, CellStyle cellStyle, CellStyleEntity style) {
        if (StringUtils.isEmpty(style.getBackground())) {
            return;
        }
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND); // 填充图案
        if (cell instanceof XSSFCell) {
            ((XSSFCellStyle) cellStyle)
                .setFillForegroundColor(PoiCssUtils.parseColor(style.getBackground()));
        } else if (cell instanceof HSSFCell) {
            cellStyle.setFillForegroundColor(
                PoiCssUtils.parseColor((HSSFWorkbook) cell.getRow().getSheet().getWorkbook(),
                    style.getBackground()).getIndex());
        }
    }

}
