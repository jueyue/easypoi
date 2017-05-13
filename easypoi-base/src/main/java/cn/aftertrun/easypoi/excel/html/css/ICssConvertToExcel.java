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
package cn.aftertrun.easypoi.excel.html.css;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

import cn.aftertrun.easypoi.excel.html.entity.style.CellStyleEntity;

/**
 * CSS Cell Style 转换类 
 * @author JueYue
 * 2016年3月20日 下午4:53:04
 */
public interface ICssConvertToExcel {
    /**
     * 把HTML样式转换成Cell样式
     * @param cell
     * @param style
     */
    public void convertToExcel(Cell cell, CellStyle cellStyle, CellStyleEntity style);

}
