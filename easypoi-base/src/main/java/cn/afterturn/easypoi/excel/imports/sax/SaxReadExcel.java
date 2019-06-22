/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package cn.afterturn.easypoi.excel.imports.sax;

import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.imports.sax.parse.ISaxRowRead;
import cn.afterturn.easypoi.excel.imports.sax.parse.SaxRowRead;
import cn.afterturn.easypoi.exception.excel.ExcelImportException;
import cn.afterturn.easypoi.handler.inter.IReadHandler;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.Iterator;
import java.util.List;

/**
 * 基于SAX Excel大数据读取,读取Excel 07版本,不支持图片读取
 *
 * @author JueYue
 * 2014年12月29日 下午9:41:38
 * @version 1.0
 */
@SuppressWarnings("rawtypes")
public class SaxReadExcel {

    private static final Logger LOGGER = LoggerFactory.getLogger(SaxReadExcel.class);

    public <T> List<T> readExcel(InputStream inputstream, Class<?> pojoClass, ImportParams params,
                                 IReadHandler hanlder) {
        try {
            OPCPackage opcPackage = OPCPackage.open(inputstream);
            return readExcel(opcPackage, pojoClass, params, null, hanlder);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelImportException(e.getMessage());
        }
    }

    private <T> List<T> readExcel(OPCPackage opcPackage, Class<?> pojoClass, ImportParams params,
                                  ISaxRowRead rowRead, IReadHandler handler) {
        try {
            XSSFReader         xssfReader         = new XSSFReader(opcPackage);
            SharedStringsTable sharedStringsTable = xssfReader.getSharedStringsTable();
            StylesTable        stylesTable        = xssfReader.getStylesTable();
            if (rowRead == null) {
                rowRead = new SaxRowRead(pojoClass, params, handler);
            }
            XMLReader             parser     = fetchSheetParser(sharedStringsTable, stylesTable, rowRead);
            Iterator<InputStream> sheets     = xssfReader.getSheetsData();
            int                   sheetIndex = 0;
            while (sheets.hasNext() && sheetIndex < params.getSheetNum() + params.getStartSheetIndex()) {
                if (sheetIndex < params.getStartSheetIndex()) {
                    sheets.next();
                } else {
                    InputStream sheet       = sheets.next();
                    InputSource sheetSource = new InputSource(sheet);
                    parser.parse(sheetSource);
                    sheet.close();
                }
                sheetIndex++;

            }
            if (handler != null) {
                handler.doAfterAll();
            }
            return rowRead.getList();
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelImportException("SAX导入数据失败");
        }
    }

    private XMLReader fetchSheetParser(SharedStringsTable sharedStringsTable, StylesTable stylesTable,
                                       ISaxRowRead rowRead) throws SAXException {
        XMLReader      parser  = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        ContentHandler handler = new SheetHandler(sharedStringsTable, stylesTable, rowRead);
        parser.setContentHandler(handler);
        return parser;
    }

}
