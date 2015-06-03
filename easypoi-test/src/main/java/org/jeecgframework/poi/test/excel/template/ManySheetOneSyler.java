package org.jeecgframework.poi.test.excel.template;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerDefaultImpl;

public class ManySheetOneSyler extends ExcelExportStylerDefaultImpl {

    private static CellStyle stringSeptailStyle;

    private static CellStyle stringNoneStyle;

    public ManySheetOneSyler(Workbook workbook) {
        super(workbook);
    }

    @Override
    public CellStyle stringSeptailStyle(Workbook workbook, boolean isWarp) {
        if (stringSeptailStyle == null) {
            stringSeptailStyle = workbook.createCellStyle();
            stringSeptailStyle.setAlignment(CellStyle.ALIGN_CENTER);
            stringSeptailStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            stringSeptailStyle.setDataFormat(STRING_FORMAT);
            stringSeptailStyle.setWrapText(true);
        }
        return stringSeptailStyle;
    }

    @Override
    public CellStyle stringNoneStyle(Workbook workbook, boolean isWarp) {
        if (stringNoneStyle == null) {
            stringNoneStyle = workbook.createCellStyle();
            stringNoneStyle.setAlignment(CellStyle.ALIGN_CENTER);
            stringNoneStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            stringNoneStyle.setDataFormat(STRING_FORMAT);
            stringNoneStyle.setWrapText(true);
        }
        return stringNoneStyle;
    }

}
