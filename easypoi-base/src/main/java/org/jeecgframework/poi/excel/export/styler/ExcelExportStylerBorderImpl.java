package org.jeecgframework.poi.excel.export.styler;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 带有边框的Excel样式
 * @author JueYue
 * @date 2015年1月9日 下午5:55:29
 */
public class ExcelExportStylerBorderImpl extends AbstractExcelExportStyler implements
                                                                         IExcelExportStyler {

    public ExcelExportStylerBorderImpl(Workbook workbook) {
        super.createStyles(workbook);
    }

    @Override
    public CellStyle getHeaderStyle(short color) {
        CellStyle titleStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 12);
        titleStyle.setFont(font);
        titleStyle.setBorderLeft((short) 1); // 左边框
        titleStyle.setBorderRight((short) 1); // 右边框
        titleStyle.setBorderBottom((short) 1);
        titleStyle.setBorderTop((short) 1);
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        return titleStyle;
    }

    @Override
    public CellStyle stringNoneStyle(Workbook workbook, boolean isWarp) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderLeft((short) 1); // 左边框
        style.setBorderRight((short) 1); // 右边框
        style.setBorderBottom((short) 1);
        style.setBorderTop((short) 1);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setDataFormat(STRING_FORMAT);
        if (isWarp) {
            style.setWrapText(true);
        }
        return style;
    }

    @Override
    public CellStyle getTitleStyle(short color) {
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setBorderLeft((short) 1); // 左边框
        titleStyle.setBorderRight((short) 1); // 右边框
        titleStyle.setBorderBottom((short) 1);
        titleStyle.setBorderTop((short) 1);
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        titleStyle.setWrapText(true);
        return titleStyle;
    }

    @Override
    public CellStyle stringSeptailStyle(Workbook workbook, boolean isWarp) {
        return isWarp ? stringNoneWrapStyle : stringNoneStyle;
    }

}
