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
package cn.aftertrun.easypoi.pdf;

import java.io.OutputStream;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import com.itextpdf.text.Document;

import cn.aftertrun.easypoi.excel.entity.params.ExcelExportEntity;
import cn.aftertrun.easypoi.pdf.entity.PdfExportParams;
import cn.aftertrun.easypoi.pdf.export.PdfExportServer;

/**
 * PDF 导出工具类
 *  
 * @author JueYue
 *  2015年10月6日 下午8:14:01
 * @version 1.0
 */
public class PdfExportUtil {

    /**
     * 根据注解导出数据
     * @param entity
     *            表格标题属性
     * @param pojoClass
     *            PDF对象Class
     * @param dataSet
     *            PDF对象数据List
     */
    public static Document exportPdf(PdfExportParams entity, Class<?> pojoClass,
                                     Collection<?> dataSet, OutputStream outStream) {
        return new PdfExportServer(outStream, entity).createPdf(entity, pojoClass, dataSet);
    }

    /**
     * 根据Map创建对应的PDF
     * @param entity
     *            表格标题属性
     * @param entityList
     *            PDF对象Class
     * @param dataSet
     *            PDF对象数据List
     */
    public static Document exportPdf(PdfExportParams entity, List<ExcelExportEntity> entityList,
                                     Collection<? extends Map<?, ?>> dataSet,
                                     OutputStream outStream) {

        return new PdfExportServer(outStream, entity).createPdfByExportEntity(entity, entityList,
            dataSet);
    }

}
