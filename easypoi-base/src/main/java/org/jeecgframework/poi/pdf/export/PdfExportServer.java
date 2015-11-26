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
package org.jeecgframework.poi.pdf.export;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.export.base.ExportBase;
import org.jeecgframework.poi.util.PoiPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Font.FontFamily;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

/**
 * PDF导出服务,基于Excel基础的导出
 * @author JueYue
 * @date 2015年10月6日 下午8:21:08
 */
public class PdfExportServer extends ExportBase {

    private static final Logger LOGGER = LoggerFactory.getLogger(PdfExportServer.class);

    private Document document;

    public PdfExportServer(OutputStream outStream) {
        try {
            document = new Document();
            PdfWriter.getInstance(document, outStream);
            document.open();
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }
    }

    public Document createPdf(ExportParams entity, Class<?> pojoClass, Collection<?> dataSet) {
        try {
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity(entity));
            }
            // 得到所有字段
            Field fileds[] = PoiPublicUtil.getClassFields(pojoClass);
            ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
            String targetId = etarget == null ? null : etarget.value();
            getAllExcelField(entity.getExclusions(), targetId, fileds, excelParams, pojoClass,
                null);
            sortAllParams(excelParams);
            int index = entity.isCreateHeadRows() ? createHeaderAndTitle(entity, excelParams) : 0;
            int titleHeight = index;
            //setCellWith(excelParams, sheet);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        } finally {
            document.close();
        }
        return document;
    }

    private int createHeaderAndTitle(ExportParams entity,
                                     List<ExcelExportEntity> excelParams) throws DocumentException {
        PdfPTable table = new PdfPTable(excelParams.size());
        for (int i = 0; i < excelParams.size(); i++) {
            table.addCell(new Phrase(excelParams.get(i).getName(), getFont()));
        }
        for (int i = 0; i < excelParams.size(); i++) {
            table.addCell(new Phrase("test", getFont()));
        }
        document.add(table);
        return 0;
    }

    private Font getFont() {
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

    public Document createPdfForMap(ExportParams entity, List<ExcelExportEntity> entityList,
                                    Collection<? extends Map<?, ?>> dataSet) {

        return document;
    }

}
