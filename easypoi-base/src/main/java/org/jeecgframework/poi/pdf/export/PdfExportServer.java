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

import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.export.base.ExportBase;
import org.jeecgframework.poi.pdf.entity.PdfExportParams;
import org.jeecgframework.poi.pdf.styler.IPdfExportStyler;
import org.jeecgframework.poi.pdf.styler.PdfExportStylerDefaultImpl;
import org.jeecgframework.poi.util.PoiPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

/**
 * PDF导出服务,基于Excel基础的导出
 * @author JueYue
 * @date 2015年10月6日 下午8:21:08
 */
public class PdfExportServer extends ExportBase {

    private static final Logger LOGGER     = LoggerFactory.getLogger(PdfExportServer.class);

    private Document            document;
    private IPdfExportStyler    styler;

    private boolean             isListData = false;

    public PdfExportServer(OutputStream outStream, PdfExportParams entity) {
        try {
            styler = entity.getStyler() == null ? new PdfExportStylerDefaultImpl()
                : entity.getStyler();
            document = styler.getDocument();
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
    public Document createPdf(PdfExportParams entity, Class<?> pojoClass, Collection<?> dataSet) {
        try {
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                //excelParams.add(indexExcelEntity(entity));
            }
            // 得到所有字段
            Field fileds[] = PoiPublicUtil.getClassFields(pojoClass);
            ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
            String targetId = etarget == null ? null : etarget.value();
            getAllExcelField(entity.getExclusions(), targetId, fileds, excelParams, pojoClass,
                null);
            createPdfByExportEntity(entity, excelParams, dataSet);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        } finally {
            document.close();
        }
        return document;
    }

    public Document createPdfByExportEntity(PdfExportParams entity,
                                            List<ExcelExportEntity> excelParams,
                                            Collection<?> dataSet) {
        try {
            sortAllParams(excelParams);
            for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
                if (excelParams.get(k).getList() != null) {
                    isListData = true;
                }
            }
            //设置各个列的宽度
            float[] widths = getCellWidths(excelParams);
            PdfPTable table = new PdfPTable(widths.length);
            table.setTotalWidth(widths);
            //table.setLockedWidth(true);
            //设置表头
            createHeaderAndTitle(entity, table, excelParams);
            int rowHeight = getRowHeight(excelParams) / 50;
            Iterator<?> its = dataSet.iterator();
            while (its.hasNext()) {
                Object t = its.next();
                createCells(table, t, excelParams, rowHeight);
            }
            document.add(table);
        } catch (DocumentException e) {
            LOGGER.error(e.getMessage(), e);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }
        return document;
    }

    private void createCells(PdfPTable table, Object t, List<ExcelExportEntity> excelParams,
                             int rowHeight) throws Exception {
        ExcelExportEntity entity;
        int maxHeight = getThisMaxHeight(t, excelParams);
        for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
            entity = excelParams.get(k);
            if (entity.getList() != null) {
                Collection<?> list = getListCellValue(entity, t);
                for (Object obj : list) {
                    createListCells(table, obj, entity.getList(), rowHeight);
                }
            } else {
                Object value = getCellValue(entity, t);
                if (entity.getType() == 1) {
                    createStringCell(table, value == null ? "" : value.toString(), entity,
                        rowHeight, 1, maxHeight);
                } else {

                }
            }
        }
    }

    /**
     * 创建集合对象
     * @param table
     * @param obj 
     * @param rowHeight 
     * @param list
     * @throws Exception 
     */
    private void createListCells(PdfPTable table, Object obj, List<ExcelExportEntity> excelParams,
                                 int rowHeight) throws Exception {
        ExcelExportEntity entity;
        for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
            entity = excelParams.get(k);
            Object value = getCellValue(entity, obj);
            if (entity.getType() == 1) {
                createStringCell(table, value == null ? "" : value.toString(), entity, rowHeight);
            } else {

            }
        }
    }

    /**
     * 获取这一列的高度
     * @param t             对象
     * @param excelParams   属性列表
     * @return
     * @throws Exception    通过反射过去值得异常
     */
    private int getThisMaxHeight(Object t, List<ExcelExportEntity> excelParams) throws Exception {
        if (isListData) {
            ExcelExportEntity entity;
            int maxHeight = 1;
            for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
                entity = excelParams.get(k);
                if (entity.getList() != null) {
                    Collection<?> list = getListCellValue(entity, t);
                    maxHeight = (list == null || maxHeight > list.size()) ? maxHeight : list.size();
                }
            }
            return maxHeight;
        }
        return 1;
    }

    /**
     * 获取Cells的宽度数组
     * @param excelParams
     * @return
     */
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

    private void createHeaderAndTitle(PdfExportParams entity, PdfPTable table,
                                      List<ExcelExportEntity> excelParams) throws DocumentException {
        int feildWidth = getFieldLength(excelParams);
        if (entity.getTitle() != null) {
            createHeaderRow(entity, table, feildWidth);
        }
        createTitleRow(entity, table, excelParams);
    }

    /**
     * 创建表头
     * 
     * @param title
     * @param index
     */
    private int createTitleRow(PdfExportParams title, PdfPTable table,
                               List<ExcelExportEntity> excelParams) {
        int rows = getRowNums(excelParams);
        for (int i = 0, exportFieldTitleSize = excelParams.size(); i < exportFieldTitleSize; i++) {
            ExcelExportEntity entity = excelParams.get(i);
            if (entity.getList() != null) {
                if (StringUtils.isNotBlank(entity.getName())) {
                    createStringCell(table, entity.getName(), entity, 10, entity.getList().size(),
                        1);
                }
                List<ExcelExportEntity> sTitel = entity.getList();
                for (int j = 0, size = sTitel.size(); j < size; j++) {
                    createStringCell(table, sTitel.get(j).getName(), sTitel.get(j), 10);
                }
            } else {
                createStringCell(table, entity.getName(), entity, 10, 1, rows == 2 ? 2 : 1);
            }
        }
        return rows;

    }

    private void createHeaderRow(PdfExportParams entity, PdfPTable table, int feildLength) {
        PdfPCell iCell = new PdfPCell(
            new Phrase(entity.getTitle(), styler.getFont(null, entity.getTitle())));
        iCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        iCell.setVerticalAlignment(Element.ALIGN_CENTER);
        iCell.setFixedHeight(entity.getTitleHeight());
        iCell.setColspan(feildLength + 1);
        table.addCell(iCell);
        if (entity.getSecondTitle() != null) {
            iCell = new PdfPCell(
                new Phrase(entity.getSecondTitle(), styler.getFont(null, entity.getSecondTitle())));
            iCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            iCell.setVerticalAlignment(Element.ALIGN_CENTER);
            iCell.setFixedHeight(entity.getSecondTitleHeight());
            iCell.setColspan(feildLength + 1);
            table.addCell(iCell);
        }
    }

    private PdfPCell createStringCell(PdfPTable table, String text, ExcelExportEntity entity,
                                      int rowHeight, int colspan, int rowspan) {
        PdfPCell iCell = new PdfPCell(new Phrase(text, styler.getFont(entity, text)));
        styler.setCellStyler(iCell, entity, text);
        iCell.setFixedHeight((int) (rowHeight * 2.5));
        if (colspan > 1) {
            iCell.setColspan(colspan);
        }
        if (rowspan > 1) {
            iCell.setRowspan(rowspan);
        }
        table.addCell(iCell);
        return iCell;
    }

    private PdfPCell createStringCell(PdfPTable table, String text, ExcelExportEntity entity,
                                      int rowHeight) {
        PdfPCell iCell = new PdfPCell(new Phrase(text, styler.getFont(entity, text)));
        styler.setCellStyler(iCell, entity, text);
        iCell.setFixedHeight((int) (rowHeight * 2.5));
        table.addCell(iCell);
        return iCell;
    }

}
