package org.jeecgframework.poi.excel.html.convert;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Excel 转换成为Html
 * @author JueYue
 * @date 2015年5月7日 下午10:06:51
 */
public final class ExcelConverHtmlUtil {

    private static final String STARTHTML = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01//EN\" \"http://www.w3.org/TR/html4/strict.dtd\"><html><head>";
    private static final String HTML_TR_S = "<tr>";
    private static final String HTML_TR_E = "</tr>";
    private static final String HTML_TD_S = "<td>";
    private static final String HTML_TD_E = "</td>";

    private ExcelConverHtmlUtil() {

    }

    public String convert(Workbook workbook, int... sheets) {
        StringBuilder html = new StringBuilder(STARTHTML);
        StringBuilder css = new StringBuilder();
        StringBuilder body = new StringBuilder();
        sheetConverToHtml(workbook.getSheetAt(sheets[0]), css, body);
        html.append(css);
        html.append(body);
        return html.toString();
    }

    private void sheetConverToHtml(Sheet sheet, StringBuilder css, StringBuilder body) {
        Iterator<Row> rows = sheet.rowIterator();
        Iterator<Cell> cells = null;
        while (rows.hasNext()) {
            Row row = rows.next();
            cells = row.cellIterator();
            body.append(HTML_TR_S);
            while (cells.hasNext()) {
                Cell cell = cells.next();
                body.append(HTML_TD_S);
                body.append(cell.toString());
                body.append(HTML_TD_E);
            }
            body.append(HTML_TR_E);
        }
    }

}
