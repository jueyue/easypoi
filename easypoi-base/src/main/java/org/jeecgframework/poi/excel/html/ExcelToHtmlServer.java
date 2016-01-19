package org.jeecgframework.poi.excel.html;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Formatter;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.entity.ExcelToHtmlParams;
import org.jeecgframework.poi.excel.html.helper.CellValueHelper;
import org.jeecgframework.poi.excel.html.helper.MergedRegionHelper;
import org.jeecgframework.poi.excel.html.helper.StylerHelper;
import org.jeecgframework.poi.util.PoiPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Excel转换成Html 服务
 * @author JueYue
 *  2015年5月10日 上午11:41:15
 */
public class ExcelToHtmlServer {

    private static final Logger           LOGGER         = LoggerFactory
        .getLogger(ExcelToHtmlServer.class);

    private static final SimpleDateFormat DATE_FORMAT    = new SimpleDateFormat("yyyy_MM_dd");

    private String                        today;

    private Workbook                      wb;
    private int                           sheetNum;
    private int                           cssRandom;

    /*是不是完成界面*/
    private boolean                       completeHTML;
    private Formatter                     out;
    /*已经完成范围处理*/
    private boolean                       gotBounds;
    private int                           firstColumn;
    private int                           endColumn;
    private String                        imageCachePath;
    private static final String           COL_HEAD_CLASS = "colHeader";
    //private static final String ROW_HEAD_CLASS = "rowHeader";

    private static final String           DEFAULTS_CLASS = "excelDefaults";

    //图片缓存
    private Map<String, PictureData>      pictures       = new HashMap<String, PictureData>();

    public ExcelToHtmlServer(ExcelToHtmlParams params) {
        this.wb = params.getWb();
        this.completeHTML = params.isCompleteHTML();
        this.sheetNum = params.getSheetNum();
        cssRandom = (int) Math.ceil(Math.random() * 1000);
        this.imageCachePath = params.getPath();
        today = DATE_FORMAT.format(new Date());
    }

    public String printPage() {
        try {
            ensureOut();
            if (completeHTML) {
                out.format("<!DOCTYPE HTML>%n");
                out.format("<html>%n");
                out.format(
                    "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\">%n");
                out.format("<head>%n");
            }
            if (StringUtils.isNotEmpty(imageCachePath)) {
                getPictures();
            }
            new StylerHelper(wb, out, sheetNum, cssRandom);
            if (completeHTML) {
                out.format("</head>%n");
                out.format("<body>%n");
            }
            print();
            if (completeHTML) {
                out.format("</body>%n");
                out.format("</html>%n");
            }
            return out.toString();
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        } finally {
            if (out != null)
                out.close();
        }
        return null;
    }

    /**
     * 获取Sheet缓存的图片
     */
    private void getPictures() {
        if (wb instanceof XSSFWorkbook) {
            pictures = PoiPublicUtil.getSheetPictrues07((XSSFSheet) wb.getSheetAt(sheetNum),
                (XSSFWorkbook) wb);
        } else {
            pictures = PoiPublicUtil.getSheetPictrues03((HSSFSheet) wb.getSheetAt(sheetNum),
                (HSSFWorkbook) wb);
        }
    }

    private void print() {
        printSheets();
    }

    private void ensureOut() {
        if (out == null)
            out = new Formatter(new StringBuilder());
    }

    private void printSheets() {
        Sheet sheet = wb.getSheetAt(sheetNum);
        printSheet(sheet);
    }

    private void printSheet(Sheet sheet) {
        out.format("<table class='%s' width='%s'>%n", DEFAULTS_CLASS, getTableWidth(sheet));
        printCols(sheet);
        printSheetContent(sheet);
        out.format("</table>%n");
    }

    private void printCols(Sheet sheet) {
        //out.format("<col/>%n");
        ensureColumnBounds(sheet);
        for (int i = firstColumn; i < endColumn; i++) {
            out.format("<col style='width:%spx;' />%n", sheet.getColumnWidth(i) / 32);
        }
    }

    private int getTableWidth(Sheet sheet) {
        ensureColumnBounds(sheet);
        int width = 0;
        for (int i = firstColumn; i < endColumn; i++) {
            width = width + (sheet.getColumnWidth(i) / 32);
        }
        return width;
    }

    private void ensureColumnBounds(Sheet sheet) {
        if (gotBounds)
            return;

        Iterator<Row> iter = sheet.rowIterator();
        firstColumn = (iter.hasNext() ? Integer.MAX_VALUE : 0);
        endColumn = 0;
        while (iter.hasNext()) {
            Row row = iter.next();
            short firstCell = row.getFirstCellNum();
            if (firstCell >= 0) {
                firstColumn = Math.min(firstColumn, firstCell);
                endColumn = Math.max(endColumn, row.getLastCellNum());
            }
        }
        gotBounds = true;
    }

    @SuppressWarnings("unused")
    /**本来是用来生成 A，B 那个列名称的**/
    private void printColumnHeads(Sheet sheet) {
        out.format("<thead>%n");
        out.format("  <tr class=%s>%n", COL_HEAD_CLASS);
        out.format("    <th class=%s>&#x25CA;</th>%n", COL_HEAD_CLASS);
        StringBuilder colName = new StringBuilder();
        for (int i = firstColumn; i < endColumn; i++) {
            colName.setLength(0);
            int cnum = i;
            do {
                colName.insert(0, (char) ('A' + cnum % 26));
                cnum /= 26;
            } while (cnum > 0);
            out.format("    <th class=%s>%s</th>%n", COL_HEAD_CLASS, colName);
        }
        out.format("  </tr>%n");
        out.format("</thead>%n");
    }

    private void printSheetContent(Sheet sheet) {
        //printColumnHeads(sheet);
        MergedRegionHelper mergedRegionHelper = new MergedRegionHelper(sheet);
        CellValueHelper cellValueHelper = new CellValueHelper(wb, cssRandom);
        out.format("<tbody>%n");
        Iterator<Row> rows = sheet.rowIterator();
        int rowIndex = 1;
        while (rows.hasNext()) {
            Row row = rows.next();
            out.format("  <tr style='height:%spx;'>%n", row.getHeight() / 15);
            //out.format("    <td class='%s'>%d</td>%n", ROW_HEAD_CLASS, row.getRowNum() + 1);
            for (int i = firstColumn; i < endColumn; i++) {
                if (mergedRegionHelper.isNeedCreate(rowIndex, i)) {
                    String content = "&nbsp;";
                    CellStyle style = null;
                    if (i >= row.getFirstCellNum() && i < row.getLastCellNum()) {
                        Cell cell = row.getCell(i);
                        if (cell != null) {
                            style = cell.getCellStyle();
                            content = cellValueHelper.getHtmlValue(cell);
                        }
                    }
                    if (pictures.containsKey((rowIndex - 1) + "_" + i)) {
                        content = "<img src='" + getImageSrc(pictures.get((rowIndex - 1) + "_" + i))
                                  + "' style='max-width:  "
                                  + getImageMaxWidth(
                                      mergedRegionHelper.getRowAndColSpan(rowIndex, i), i, sheet)
                                  + "px;' />";
                    }
                    if (mergedRegionHelper.isMergedRegion(rowIndex, i)) {
                        Integer[] rowAndColSpan = mergedRegionHelper.getRowAndColSpan(rowIndex, i);
                        out.format("    <td rowspan='%s' colspan='%s' class='%s' >%s</td>%n",
                            rowAndColSpan[0], rowAndColSpan[1], styleName(style), content);
                    } else {
                        out.format("    <td class='%s'>%s</td>%n", styleName(style), content);
                    }
                }

            }
            out.format("  </tr>%n");
            rowIndex++;
        }
        out.format("</tbody>%n");
    }

    /**
     * 获取图片最大宽度
     * @param colIndex 
     * @param sheet 
     * @param span 
     * @return
     */
    private int getImageMaxWidth(Integer[] rowAndColSpan, int colIndex, Sheet sheet) {
        if (rowAndColSpan == null) {
            return sheet.getColumnWidth(colIndex) / 32;
        }
        int maxWidth = 0;
        for (int i = 0; i < rowAndColSpan[1]; i++) {
            maxWidth += sheet.getColumnWidth(colIndex + i) / 32;
        }
        return maxWidth;
    }

    /**
     * 获取图片路径
     * @param pictureData
     * @return
     */
    private String getImageSrc(PictureData pictureData) {
        if (pictureData == null) {
            return "";
        }
        byte[] data = pictureData.getData();
        String fileName = "pic" + Math.round(Math.random() * 100000000000L);
        fileName += "." + PoiPublicUtil.getFileExtendName(data);
        if (!imageCachePath.startsWith("/") && !imageCachePath.contains(":")) {
            imageCachePath = PoiPublicUtil.getWebRootPath(imageCachePath);
        }
        File savefile = new File(imageCachePath + "/" + today);
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        savefile = new File(imageCachePath + "/" + today + "/" + fileName);
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(savefile);
            fos.write(data);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        } finally {
            try {
                fos.close();
            } catch (Exception e) {
                LOGGER.error(e.getMessage(), e);
            }
        }
        return imageCachePath + "/" + today + "/" + fileName;
    }

    private String styleName(CellStyle style) {
        if (style == null)
            return "";
        return String.format("style_%02x_%s font_%s_%s", style.getIndex(), cssRandom,
            style.getFontIndex(), cssRandom);
    }
}
