package org.jeecgframework.poi.excel.export.styler;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 抽象接口提供两个公共方法
 * @author JueYue
 * @date 2015年1月9日 下午5:48:55
 */
public abstract class AbstractExcelExportStyler implements IExcelExportStyler {

    protected CellStyle          oneStyle;
    protected CellStyle          oneWrapStyle;
    protected CellStyle          doubleStyle;
    protected CellStyle          doubleWrapStyle;
    protected Workbook           workbook;

    protected static final short cellFormat = (short) BuiltinFormats.getBuiltinFormat("TEXT");

    protected void createStyles(Workbook workbook) {
        this.oneStyle = createOneStyle(workbook, false);
        this.oneWrapStyle = createOneStyle(workbook, true);
        this.doubleStyle = createDoubleStyle(workbook, false);
        this.doubleWrapStyle = createDoubleStyle(workbook, true);
        this.workbook = workbook;
    }

    @Override
    public CellStyle getStyles(boolean needOne, boolean isWrap) {
        if (needOne && isWrap) {
            return oneWrapStyle;
        }
        if (needOne) {
            return oneStyle;
        }
        if (needOne == false && isWrap) {
            return doubleWrapStyle;
        }
        return doubleStyle;
    }

    @Override
    public CellStyle createOneStyle(Workbook workbook, boolean isWarp) {
        return null;
    }

    @Override
    public CellStyle createDoubleStyle(Workbook workbook, boolean isWarp) {
        return null;
    }

}
