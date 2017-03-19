/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 *   
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.jeecgframework.poi.excel.html.css.impl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.jeecgframework.poi.excel.html.css.ICssConvertToExcel;
import org.jeecgframework.poi.excel.html.css.ICssConvertToHtml;
import org.jeecgframework.poi.excel.html.entity.CellStyleBorderEntity;
import org.jeecgframework.poi.excel.html.entity.CellStyleEntity;

/**
 * 边框转换实现类
 * @author JueYue
 * 2016年4月3日 上午10:26:47
 */
public class BorderCssConverImpl implements ICssConvertToExcel, ICssConvertToHtml {

    @Override
    public String convertToHtml(Cell cell, CellStyle cellStyle, CellStyleEntity style) {
        CellStyleBorderEntity border = new CellStyleBorderEntity();
        border.setBorderBottom(cellStyle.getBorderBottom());
        border.setBorderBottomColor(cellStyle.getBottomBorderColor());
        border.setBorderLeft(cellStyle.getBorderLeft());
        border.setBorderLeftColor(cellStyle.getLeftBorderColor());
        border.setBorderRight(cellStyle.getBorderRight());
        border.setBorderRightColor(cellStyle.getRightBorderColor());
        border.setBorderTop(cellStyle.getBorderTop());
        border.setBorderTopColor(cellStyle.getTopBorderColor());
        style.setBorder(border);
        return null;
    }

    @Override
    public void convertToExcel(Cell cell, CellStyle cellStyle, CellStyleEntity style) {
    }

}
