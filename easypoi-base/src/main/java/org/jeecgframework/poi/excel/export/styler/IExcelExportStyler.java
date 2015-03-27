package org.jeecgframework.poi.excel.export.styler;

import org.apache.poi.ss.usermodel.CellStyle;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;

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
     * @param noneStyler
     * @param entity
     * @return
     */
    public CellStyle getStyles(boolean noneStyler, ExcelExportEntity entity);

}
