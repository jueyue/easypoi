package org.jeecgframework.poi.excel.imports.sax;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.jeecgframework.poi.excel.imports.sax.parse.ISaxRowRead;
import org.jeecgframework.poi.excel.imports.sax.parse.SaxRowRead;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * 基于SAX Excel大数据读取,读取Excel 07版本,不支持图片读取
 * @author JueYue
 * @date 2014年12月29日 下午9:41:38
 * @version 1.0
 */
public class SaxReadExcel {

    public <T> List<T> readExcel(InputStream inputstream, Class<?> pojoClass, ImportParams params) {
        try {
            OPCPackage opcPackage = OPCPackage.open(inputstream);
            XSSFReader xssfReader = new XSSFReader(opcPackage);
            SharedStringsTable sst = xssfReader.getSharedStringsTable();
            ISaxRowRead rowRead = new SaxRowRead();
            XMLReader parser = fetchSheetParser(sst, rowRead);
            Iterator<InputStream> sheets = xssfReader.getSheetsData();
            int sheetIndex = 0;
            while (sheets.hasNext() && sheetIndex < params.getSheetNum()) {
                sheetIndex++;
                InputStream sheet = sheets.next();
                InputSource sheetSource = new InputSource(sheet);
                parser.parse(sheetSource);
                sheet.close();
            }
            return rowRead.getList();
        } catch (Exception e) {
        }

        return null;
    }

    public XMLReader fetchSheetParser(SharedStringsTable sst, ISaxRowRead rowRead)
                                                                                  throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        ContentHandler handler = new SheetHandler(sst, rowRead);
        parser.setContentHandler(handler);
        return parser;
    }

}
