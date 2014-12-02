package org.jeecgframework.poi.excel;

import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.export.ExcelExportServer;
/**
 * 自己定义导出的样式
 * @author JueYue
 * @date 2014年12月2日 上午10:32:44
 */
public class MyExcelExport extends ExcelExportServer {

    @Override
    public CellStyle getHeaderStyle(Workbook workbook, ExportParams entity) {
        return super.getHeaderStyle(workbook, entity);
    }

    @Override
    public CellStyle getOneStyle(Workbook workbook, boolean isWarp) {
        return super.getOneStyle(workbook, isWarp);
    }

    @Override
    public CellStyle getStyles(Map<String, CellStyle> map, boolean needOne, boolean isWrap) {
        return super.getStyles(map, needOne, isWrap);
    }

    @Override
    public CellStyle getTitleStyle(Workbook workbook, ExportParams entity) {
        return super.getTitleStyle(workbook, entity);
    }

    @Override
    public CellStyle getTwoStyle(Workbook workbook, boolean isWarp) {
        return super.getTwoStyle(workbook, isWarp);
    }

}
