package org.jeecgframework.poi.excel.export.styler;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Excel导出样式接口
 * @author JueYue
 * @date 2015年1月9日 下午5:32:30
 */
public interface IExcelExportStyler {

    /**
     * 列表头样式
     * @param headerColor
     * @return
     */
    public CellStyle getHeaderStyle(short headerColor);

    /**
     * 标题样式
     * @param color
     * @return
     */
    public CellStyle getTitleStyle(short color);

    /**
     * 获取样式方法
     * @param map
     * @param needOne
     * @param isWrap
     * @return
     */
    public CellStyle getStyles(boolean needOne, boolean isWrap);

    /**
     * 创建单行样式
     * @param workbook
     * @param isWarp
     * @return
     */
    public CellStyle createOneStyle(Workbook workbook, boolean isWarp);

    /**
     * 创建双行样式
     * @param workbook
     * @param isWarp
     * @return
     */
    public CellStyle createDoubleStyle(Workbook workbook, boolean isWarp);

}
