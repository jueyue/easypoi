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

import java.io.IOException;

import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Font.FontFamily;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;

/**
 * 默认的PDFstyler 实现
 * @author JueYue
 * @date 2016年1月8日 下午2:06:26
 */
public class PdfExportStylerDefaultImpl implements IPdfExportStyler {

    private static final Logger LOGGER = LoggerFactory.getLogger(PdfExportStylerDefaultImpl.class);

    @Override
    public Document getDocument() {
        return new Document(PageSize.A4, 36, 36, 24, 36);
    }

    @Override
    public void setCellStyler(PdfPCell iCell, ExcelExportEntity entity, String text) {
        iCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        iCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
    }

    @Override
    public Font getFont(ExcelExportEntity entity, String text) {
        try {
            //用以支持中文
            BaseFont bfChinese = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H",
                BaseFont.NOT_EMBEDDED);
            Font font = new Font(bfChinese);
            return font;
        } catch (DocumentException e) {
            LOGGER.error(e.getMessage(), e);
        } catch (IOException e) {
            LOGGER.error(e.getMessage(), e);
        }
        Font font = new Font(FontFamily.UNDEFINED);
        return font;
    }

}
