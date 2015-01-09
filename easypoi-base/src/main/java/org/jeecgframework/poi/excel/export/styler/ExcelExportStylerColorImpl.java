package org.jeecgframework.poi.excel.export.styler;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 带有样式的导出服务
 * @author JueYue
 * @date 2015年1月9日 下午4:54:15
 */
public class ExcelExportStylerColorImpl extends AbstractExcelExportStyler implements
                                                                        IExcelExportStyler {

    public ExcelExportStylerColorImpl(Workbook workbook) {
        super.createStyles(workbook);
    }

    @Override
    public CellStyle getHeaderStyle(short headerColor) {
        CellStyle titleStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 24);
        titleStyle.setFont(font);
        titleStyle.setFillForegroundColor(headerColor);
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        return titleStyle;
    }

    @Override
    public CellStyle createOneStyle(Workbook workbook, boolean isWarp) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderLeft((short) 1); // 左边框
        style.setBorderRight((short) 1); // 右边框
        style.setBorderBottom((short) 1);
        style.setBorderTop((short) 1);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setDataFormat(cellFormat);
        if (isWarp) {
            style.setWrapText(true);
        }
        return style;
    }

    @Override
    public CellStyle getTitleStyle(short color) {
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFillForegroundColor(color); // 填充的背景颜色
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND); // 填充图案
        titleStyle.setWrapText(true);
        return titleStyle;
    }

    @Override
    public CellStyle createDoubleStyle(Workbook workbook, boolean isWarp) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderLeft((short) 1); // 左边框
        style.setBorderRight((short) 1); // 右边框
        style.setBorderBottom((short) 1);
        style.setBorderTop((short) 1);
        style.setFillForegroundColor((short) 41); // 填充的背景颜色
        style.setFillPattern(CellStyle.SOLID_FOREGROUND); // 填充图案
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setDataFormat(cellFormat);
        if (isWarp) {
            style.setWrapText(true);
        }
        return style;
    }

}
