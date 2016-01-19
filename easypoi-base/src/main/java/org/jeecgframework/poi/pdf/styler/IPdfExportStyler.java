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
package org.jeecgframework.poi.pdf.styler;

import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;

import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.PdfPCell;

/**
 * PDF导出样式设置
 * @author JueYue
 *   2016年1月7日 下午11:16:51
 */
public interface IPdfExportStyler {

    /**
     * 获取文档格式
     * @return
     */
    public Document getDocument();

    /**
     * 设置Cell的样式
     * @param entity
     * @param text
     */
    public void setCellStyler(PdfPCell iCell, ExcelExportEntity entity, String text);

    /**
     * 获取字体
     * @param entity
     * @param text
     */
    public Font getFont(ExcelExportEntity entity, String text);

}
