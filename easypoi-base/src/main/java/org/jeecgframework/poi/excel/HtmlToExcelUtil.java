/**
 * Copyright 2013-2017 JueYue (qrb.jueyue@gmail.com)
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
package org.jeecgframework.poi.excel;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.excel.html.HtmlToExcelServer;
import org.jeecgframework.poi.exception.excel.ExcelExportException;
import org.jeecgframework.poi.exception.excel.enums.ExcelExportEnum;

/**
 * 基于Excel和Html的互换
 * @author JueYue
 * 2017年3月27日
 */
public final class HtmlToExcelUtil {

    private HtmlToExcelUtil() {

    }

    public static Workbook htmlToExcel(String html, ExcelType type) {
        Workbook workbook = null;
        if (ExcelType.HSSF.equals(type)) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }
        new HtmlToExcelServer().createSheet(html, workbook);
        return workbook;
    }

    public static Workbook htmlToExcel(InputStream is, ExcelType type) {
        try {
            return htmlToExcel(new String(IOUtils.toByteArray(is)), type);
        } catch (IOException e) {
            throw new ExcelExportException(ExcelExportEnum.HTML_ERROR, e);
        }
    }

}
