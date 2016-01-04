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
import java.util.Iterator;
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

    private Document            document;

    public PdfExportServer(OutputStream outStream) {
        try {
            document = new Document();
            PdfWriter.getInstance(document, outStream);
            document.open();
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }
    }

    /**
     * 创建Pdf的表格数据
     * @param entity
     * @param pojoClass
     * @param dataSet
     * @return
     */
    public Document createTable(ExportParams entity, Class<?> pojoClass, Collection<?> dataSet) {
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
            //设置各个列的宽度
            float[] widths = getCellWidths(excelParams);
            PdfPTable table = new PdfPTable(widths.length);
            table.setTotalWidth(widths);
            //table.setLockedWidth(true);
            //设置表头
            createHeaderAndTitle(entity, table, excelParams);
            short rowHeight = getRowHeight(excelParams);
            Iterator<?> its = dataSet.iterator();
            while (its.hasNext()) {
                Object t = its.next();
                createCells(table, t, excelParams, rowHeight);
            }
            document.add(table);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        } finally {
            document.close();
        }
        return document;
    }

    private void createCells(PdfPTable table, Object t, List<ExcelExportEntity> excelParams,
                             short rowHeight) throws Exception {
        ExcelExportEntity entity;
        int maxHeight = 1, cellNum = 0;
        //int indexKey = createIndexCell(row, index, excelParams.get(0));
        //cellNum += indexKey;
        for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
            entity = excelParams.get(k);
            if (entity.getList() != null) {
                Collection<?> list = getListCellValue(entity, t);
                int listC = 0;
                for (Object obj : list) {
                    //createListCells(patriarch, index + listC, cellNum, obj, entity.getList(), sheet,
                    //   workbook);
                    listC++;
                }
                cellNum += entity.getList().size();
                if (list != null && list.size() > maxHeight) {
                    maxHeight = list.size();
                }
            } else {
                Object value = getCellValue(entity, t);
                if (entity.getType() == 1) {
                    table.addCell(value == null ? "" : value.toString());
                } else {
                }
            }
        }
    }

    private float[] getCellWidths(List<ExcelExportEntity> excelParams) {
        List<Float> widths = new ArrayList<Float>();
        for (int i = 0; i < excelParams.size(); i++) {
            if (excelParams.get(i).getList() != null) {
                List<ExcelExportEntity> list = excelParams.get(i).getList();
                for (int j = 0; j < list.size(); j++) {
                    widths.add((float) (20 * list.get(j).getWidth()));
                }
            } else {
                widths.add((float) (20 * excelParams.get(i).getWidth()));
            }
        }
        float[] widthArr = new float[widths.size()];
        for (int i = 0; i < widthArr.length; i++) {
            widthArr[i] = widths.get(i);
        }
        return widthArr;
    }

    private void createHeaderAndTitle(ExportParams entity, PdfPTable table,
                                      List<ExcelExportEntity> excelParams) throws DocumentException {
        for (int i = 0; i < excelParams.size(); i++) {
            table.addCell(new Phrase(excelParams.get(i).getName(), getFont()));
        }
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
