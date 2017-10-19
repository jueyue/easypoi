/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package cn.afterturn.easypoi.excel.export.styler;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.entity.params.ExcelForEachParams;

/**
 * Excel导出样式接口
 *
 * @author JueYue 2015年1月9日 下午5:32:30
 */
public interface IExcelExportStyler {

    /**
     * 列表头样式
     */
    public CellStyle getHeaderStyle(short headerColor);

    /**
     * 标题样式
     */
    public CellStyle getTitleStyle(short color);

    /**
     * 获取样式方法
     */
    @Deprecated
    public CellStyle getStyles(boolean parity, ExcelExportEntity entity);

    /**
     * 获取样式方法
     *
     * @param dataRow 数据行
     * @param obj     对象
     * @param data    数据
     */
    public CellStyle getStyles(Cell cell, int dataRow, ExcelExportEntity entity, Object obj, Object data);

    /**
     * 模板使用的样式设置
     */
    public CellStyle getTemplateStyles(boolean isSingle, ExcelForEachParams excelForEachParams);

}
