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

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.jeecgframework.poi.excel.html.css.ICssConvertToExcel;
import org.jeecgframework.poi.excel.html.css.ICssConvertToHtml;
import org.jeecgframework.poi.excel.html.entity.CellStyleEntity;
import org.jeecgframework.poi.util.PoiCssUtils;


/**
 * 行高转换实现类
 * @author JueYue
 * 2016年4月3日 上午10:26:47
 */
public class HeightCssConverImpl implements ICssConvertToExcel, ICssConvertToHtml {

    @Override
    public String convertToHtml(Cell cell, CellStyle cellStyle, CellStyleEntity style) {

        return null;
    }

    @Override
    public void convertToExcel(Cell cell, CellStyle cellStyle, CellStyleEntity style) {
        if (StringUtils.isNoneBlank(style.getHeight())) {
            int height = Math.round(PoiCssUtils.getInt(style.getHeight()) * 255 / 12.75F);
            Row row = cell.getRow();
            if (height > row.getHeight()) {
                row.setHeight((short) height);
            }
        }
    }

}
