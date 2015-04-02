package org.jeecgframework.poi.test.excel.styler;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerDefaultImpl;

/**
 * Excel 自定义styler 的例子
 * @author JueYue
 * @date 2015年3月29日 下午9:04:41
 */
public class ExcelExportStatisticStyler extends ExcelExportStylerDefaultImpl {

    private CellStyle numberCellStyle;

    public ExcelExportStatisticStyler(Workbook workbook) {
        super(workbook);
        createNumberCellStyler();
    }

    private void createNumberCellStyler() {
        numberCellStyle = workbook.createCellStyle();
        numberCellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        numberCellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        numberCellStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("0.00"));
        numberCellStyle.setWrapText(true);
    }

    @Override
    public CellStyle getStyles(boolean noneStyler, ExcelExportEntity entity) {
        if (entity != null
            && (entity.getName().contains("int") || entity.getName().contains("double"))) {
            return numberCellStyle;
        }
        return super.getStyles(noneStyler, entity);
    }

}
