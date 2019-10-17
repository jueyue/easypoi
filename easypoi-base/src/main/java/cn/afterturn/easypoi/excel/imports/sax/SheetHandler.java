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

import cn.afterturn.easypoi.excel.entity.enmus.CellValueType;
import cn.afterturn.easypoi.excel.entity.sax.SaxReadCellEntity;
import cn.afterturn.easypoi.excel.imports.sax.parse.ISaxRowRead;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static cn.afterturn.easypoi.excel.entity.sax.SaxConstant.*;

/**
 * 回调接口
 *
 * @author JueYue
 * 2014年12月29日 下午9:50:09
 */
public class SheetHandler extends DefaultHandler {

    private SharedStringsTable sharedStringsTable;
    private StylesTable        stylesTable;
    private String             lastContents;

    /**
     * 当前行
     **/
    private int curRow = 0;
    /**
     * 当前列
     **/
    private int curCol = 0;

    private CellValueType type;
    private String        currentLocation, prevLocation;

    private ISaxRowRead read;

    private List<SaxReadCellEntity> rowList = new ArrayList<>();

    public SheetHandler(SharedStringsTable sharedStringsTable, StylesTable stylesTable, ISaxRowRead rowRead) {
        this.sharedStringsTable = sharedStringsTable;
        this.stylesTable = stylesTable;
        this.read = rowRead;
    }

    @Override
    public void startElement(String uri, String localName, String name,
                             Attributes attributes) throws SAXException {
        // 置空  
        lastContents = "";
        if (COL.equals(name)) {
            String cellType = attributes.getValue(TYPE);
            prevLocation = currentLocation;
            currentLocation = attributes.getValue(ROW_COL);
            if (STRING.equals(cellType)) {
                type = CellValueType.String;
                return;
            }
            if (BOOLEAN.equals(cellType)) {
                type = CellValueType.Boolean;
                return;
            }
            if (DATE.equals(cellType)) {
                type = CellValueType.Date;
                return;
            }
            if (INLINE_STR.equals(cellType)) {
                type = CellValueType.InlineStr;
                return;
            }
            if (FORMULA.equals(cellType)) {
                type = CellValueType.Formula;
                return;
            }
            if (NUMBER.equals(cellType)) {
                type = CellValueType.Number;
                return;
            }
            try {
                short  nfId         = (short) stylesTable.getCellXfAt(Integer.parseInt(attributes.getValue(STYLE))).getNumFmtId();
                String numberFormat = stylesTable.getNumberFormats().get(nfId).toUpperCase();
                if (StringUtils.isNotEmpty(numberFormat)) {
                    if (numberFormat.contains("Y") || numberFormat.contains("M") || numberFormat.contains("D")
                            || numberFormat.contains("H") || numberFormat.contains("S") || numberFormat.contains("年")
                            || numberFormat.contains("月") || numberFormat.contains("日") || numberFormat.contains("时")
                            || numberFormat.contains("分") || numberFormat.contains("秒")) {
                        type = CellValueType.Date;
                        return;
                    }
                }
            } catch (Exception e) {

            }
            // 没别的了就是数字了
            type = CellValueType.Number;
        } else if (T_ELEMENT.equals(name)) {
            type = CellValueType.TElement;
        }

    }

    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {

        // 根据SST的索引值的到单元格的真正要存储的字符串  
        // 这时characters()方法可能会被调用多次  
        if (CellValueType.String.equals(type)) {
            try {
                int idx = Integer.parseInt(lastContents);
                lastContents = sharedStringsTable.getItemAt(idx).getString();
            } catch (Exception e) {
            }
        }
        if (VALUE.equals(name) && StringUtils.isNotEmpty(prevLocation)) {
            addNullCell(prevLocation, currentLocation);
        }
        //t元素也包含字符串
        if (CellValueType.TElement.equals(type)) {
            String value = lastContents.trim();
            rowList.add(curCol, new SaxReadCellEntity(CellValueType.String, value));
            curCol++;
            type = CellValueType.None;
            // v => 单元格的值，如果单元格是字符串则v标签的值为该字符串在SST中的索引  
            // 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符  
        } else if (VALUE.equals(name)) {
            String value = lastContents.trim();
            value = "".equals(value) ? " " : value;
            if (CellValueType.Date.equals(type)) {
                Date date = HSSFDateUtil.getJavaDate(Double.valueOf(value));
                rowList.add(curCol, new SaxReadCellEntity(CellValueType.Date, date));
            } else if (CellValueType.Number.equals(type)) {
                BigDecimal bd = new BigDecimal(value);
                rowList.add(curCol, new SaxReadCellEntity(CellValueType.Number, bd));
            } else if (CellValueType.String.equals(type) || CellValueType.InlineStr.equals(type)) {
                rowList.add(curCol, new SaxReadCellEntity(CellValueType.String, value));
            }
            curCol++;
            //如果标签名称为 row ，这说明已到行尾，调用 optRows() 方法
        } else if (COL.equals(name) && StringUtils.isEmpty(lastContents)) {
            rowList.add(curCol, new SaxReadCellEntity(CellValueType.String, ""));
            curCol++;
        } else if (ROW.equals(name)) {
            read.parse(curRow, rowList);
            rowList.clear();
            curRow++;
            curCol = 0;
        }

    }

    private void addNullCell(String prevLocation, String currentLocation) {
        // 拆分行和列
        String[] prev    = getRowCell(prevLocation);
        String[] current = getRowCell(currentLocation);
        if (prev[1].equalsIgnoreCase(current[1])) {
            int prevCell    = getCellNum(prev[0]) + 1;
            int currentCell = getCellNum(current[0]);
            for (int i = prevCell; i < currentCell; i++) {
                rowList.add(curCol, new SaxReadCellEntity(CellValueType.String, ""));
                curCol++;
            }
        }
    }

    private int getCellNum(String cell) {
        if (StringUtils.isEmpty(cell)) {
            return 0;
        }
        char[] chars = cell.toUpperCase().toCharArray();
        int    n     = 0;
        for (int i = cell.length() - 1, j = 1; i >= 0; i--, j *= 26) {
            char c = (chars[i]);
            if (c < 'A' || c > 'Z') {
                return 0;
            }
            n += ((int) c - 64) * j;
        }
        return n;
    }

    private String[] getRowCell(String prevLocation) {
        StringBuilder row   = new StringBuilder();
        StringBuilder cell  = new StringBuilder();
        char[]        chars = prevLocation.toCharArray();
        for (int i = 0; i < chars.length; i++) {
            if (chars[i] >= '0' && chars[i] <= '9') {
                cell.append(chars[i]);
            } else {
                row.append(chars[i]);
            }
        }
        return new String[]{row.toString(), cell.toString()};
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        //得到单元格内容的值  
        lastContents += new String(ch, start, length);
    }

}
