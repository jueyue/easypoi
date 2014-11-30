package org.jeecgframework.poi.excel;

import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.export.ExcelExportServer;

public class MyExcelExport extends ExcelExportServer {

    @Override
    public HSSFCellStyle getHeaderStyle(HSSFWorkbook workbook, ExportParams entity) {
        return super.getHeaderStyle(workbook, entity);
    }

    @Override
    public HSSFCellStyle getOneStyle(HSSFWorkbook workbook, boolean isWarp) {
        return super.getOneStyle(workbook, isWarp);
    }

    @Override
    public CellStyle getStyles(Map<String, HSSFCellStyle> map, boolean needOne, boolean isWrap) {
        return super.getStyles(map, needOne, isWrap);
    }

    @Override
    public HSSFCellStyle getTitleStyle(HSSFWorkbook workbook, ExportParams entity) {
        return super.getTitleStyle(workbook, entity);
    }

    @Override
    public HSSFCellStyle getTwoStyle(HSSFWorkbook workbook, boolean isWarp) {
        return super.getTwoStyle(workbook, isWarp);
    }

}
