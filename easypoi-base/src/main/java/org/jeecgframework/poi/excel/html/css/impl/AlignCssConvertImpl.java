package org.jeecgframework.poi.excel.html.css.impl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.jeecgframework.poi.excel.html.css.ICssConvertToExcel;
import org.jeecgframework.poi.excel.html.css.ICssConvertToHtml;
import org.jeecgframework.poi.excel.html.entity.CellStyleEntity;

public class AlignCssConvertImpl implements ICssConvertToExcel, ICssConvertToHtml {

    @Override
    public String convertToHtml(Cell cell, CellStyle cellStyle, CellStyleEntity style) {
        
        return null;
    }

    @Override
    public void convertToExcel(Cell cell, CellStyle cellStyle, CellStyleEntity style) {
    }

   

}
