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
package cn.afterturn.easypoi.word.parse;

import cn.afterturn.easypoi.cache.WordCache;
import cn.afterturn.easypoi.entity.ImageEntity;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import cn.afterturn.easypoi.word.entity.MyXWPFDocument;
import cn.afterturn.easypoi.word.entity.params.ExcelListEntity;
import cn.afterturn.easypoi.word.parse.excel.ExcelEntityParse;
import cn.afterturn.easypoi.word.parse.excel.ExcelMapParse;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;

import static cn.afterturn.easypoi.util.PoiElUtil.*;

/**
 * 解析07版的Word,替换文字,生成表格,生成图片
 *
 * @author JueYue
 * 2013-11-16
 * @version 1.0
 */
@SuppressWarnings({"unchecked", "rawtypes"})
public class ParseWord07 {

    private static final Logger LOGGER = LoggerFactory.getLogger(ParseWord07.class);

    /**
     * 根据条件改变值
     *
     * @param map
     * @author JueYue
     * 2013-11-16
     */
    private void changeValues(XWPFParagraph paragraph, XWPFRun currentRun, String currentText,
                              List<Integer> runIndex, Map<String, Object> map) throws Exception {
        Object obj = PoiPublicUtil.getRealValue(currentText, map);
        // 如果是图片就设置为图片
        if (obj instanceof ImageEntity) {
            currentRun.setText("", 0);
            ExcelMapParse.addAnImage((ImageEntity) obj, currentRun);
        } else {
            currentText = obj.toString();
            PoiPublicUtil.setWordText(currentRun, currentText);
        }
        for (int k = 0; k < runIndex.size(); k++) {
            paragraph.getRuns().get(runIndex.get(k)).setText("", 0);
        }
        runIndex.clear();
    }

    /**
     * 判断是不是迭代输出
     *
     * @return
     * @throws Exception
     * @author JueYue
     * 2013-11-18
     */
    private Object checkThisTableIsNeedIterator(XWPFTableCell cell,
                                                Map<String, Object> map) throws Exception {
        String text = cell.getText().trim();
        // 判断是不是迭代输出
        if (text != null && text.contains(FOREACH) && text.startsWith(START_STR)) {
            text = text.replace(FOREACH_NOT_CREATE, EMPTY).replace(FOREACH_AND_SHIFT, EMPTY)
                    .replace(FOREACH, EMPTY).replace(START_STR, EMPTY);
            String[] keys   = text.replaceAll("\\s{1,}", " ").trim().split(" ");
            Object   result = PoiPublicUtil.getParamsValue(keys[0], map);
            //添加list默认值，避免将{{$fe: list t.sn	t.hoby	t.remark}} 这类标签直接显示出来
            return Objects.nonNull(result) ? result : new ArrayList<Map<String, Object>>(0);
        }
        return null;
    }

    /**
     * 解析所有的文本
     *
     * @param paragraphs
     * @param map
     * @author JueYue
     * 2013-11-17
     */
    private void parseAllParagraphic(List<XWPFParagraph> paragraphs,
                                     Map<String, Object> map) throws Exception {
        XWPFParagraph paragraph;
        for (int i = 0; i < paragraphs.size(); i++) {
            paragraph = paragraphs.get(i);
            if (paragraph.getText().indexOf(START_STR) != -1) {
                parseThisParagraph(paragraph, map);
            }

        }

    }

    /**
     * 解析这个段落
     *
     * @param paragraph
     * @param map
     * @author JueYue
     * 2013-11-16
     */
    private void parseThisParagraph(XWPFParagraph paragraph, Map<String, Object> map) throws Exception {
        XWPFRun run;
        // 拿到的第一个run,用来set值,可以保存格式
        XWPFRun currentRun = null;
        // 存放当前的text
        String currentText = "";
        String text;
        // 判断是不是已经遇到{{
        Boolean isfinde = false;
        // 存储遇到的run,把他们置空
        List<Integer> runIndex = new ArrayList<Integer>();
        for (int i = 0; i < paragraph.getRuns().size(); i++) {
            run = paragraph.getRuns().get(i);
            text = run.getText(0);
            if (StringUtils.isEmpty(text)) {
                continue;
            } // 如果为空或者""这种这继续循环跳过
            if (isfinde) {
                currentText += text;
                if (currentText.indexOf(START_STR) == -1) {
                    isfinde = false;
                    runIndex.clear();
                } else {
                    runIndex.add(i);
                }
                if (currentText.indexOf(END_STR) != -1) {
                    changeValues(paragraph, currentRun, currentText, runIndex, map);
                    currentText = "";
                    isfinde = false;
                }
                // 判断是不是开始
            } else if (text.indexOf(START_STR) >= 0) {
                currentText = text;
                isfinde = true;
                currentRun = run;
            } else {
                currentText = "";
            }
            if (currentText.indexOf(END_STR) != -1) {
                changeValues(paragraph, currentRun, currentText, runIndex, map);
                isfinde = false;
            }
        }

    }

    private void parseThisRow(List<XWPFTableCell> cells, Map<String, Object> map) throws Exception {
        for (XWPFTableCell cell : cells) {
            parseAllParagraphic(cell.getParagraphs(), map);
        }
    }


    /**
     * 解析这个表格
     *
     * @param table
     * @param map
     * @author JueYue
     * 2013-11-17
     */
    private void parseThisTable(XWPFTable table, Map<String, Object> map) throws Exception {
        XWPFTableRow        row;
        List<XWPFTableCell> cells;
        Object              listobj;
        for (int i = 0; i < table.getNumberOfRows(); i++) {
            row = table.getRow(i);
            cells = row.getTableCells();
            listobj = checkThisTableIsNeedIterator(cells.get(0), map);
            if (listobj == null) {
                parseThisRow(cells, map);
            } else if (listobj instanceof ExcelListEntity) {
                new ExcelEntityParse().parseNextRowAndAddRow(table, i, (ExcelListEntity) listobj);
                //删除之后要往上挪一行,然后加上跳过新建的行数
                i = i + ((ExcelListEntity) listobj).getList().size() - 1;
            } else {
                ExcelMapParse.parseNextRowAndAddRow(table, i, (List) listobj);
                //删除之后要往上挪一行,然后加上跳过新建的行数
                i = i + ((List) listobj).size() - 1;
            }
        }
    }

    /**
     * 解析07版的Word并且进行赋值
     *
     * @return
     * @throws Exception
     * @author JueYue
     * 2013-11-16
     */
    public XWPFDocument parseWord(String url, Map<String, Object> map) throws Exception {
        MyXWPFDocument doc = WordCache.getXWPFDocument(url);
        parseWordSetValue(doc, map);
        return doc;
    }

    /**
     * 解析07版的Work并且进行赋值但是进行多页拼接
     *
     * @param url
     * @param list
     * @return
     */
    public XWPFDocument parseWord(String url, List<Map<String, Object>> list) throws Exception {
        if (list == null || list.size() == 0) {
            return null;
        } else if (list.size() == 1) {
            return parseWord(url, list.get(0));
        } else {
            MyXWPFDocument doc = WordCache.getXWPFDocument(url);
            parseWordSetValue(doc, list.get(0));
            //插入分页
            doc.createParagraph().setPageBreak(true);
            for (int i = 1; i < list.size(); i++) {
                MyXWPFDocument tempDoc = WordCache.getXWPFDocument(url);
                parseWordSetValue(tempDoc, list.get(i));
                tempDoc.createParagraph().setPageBreak(true);
                doc.getDocument().addNewBody().set(tempDoc.getDocument().getBody());

            }
            return doc;
        }

    }

    /**
     * 解析07版的Word并且进行赋值
     *
     * @throws Exception
     */
    public void parseWord(XWPFDocument document, Map<String, Object> map) throws Exception {
        parseWordSetValue((MyXWPFDocument) document, map);
    }

    private void parseWordSetValue(MyXWPFDocument doc, Map<String, Object> map) throws Exception {
        // 第一步解析文档
        parseAllParagraphic(doc.getParagraphs(), map);
        // 第二步解析页眉,页脚
        parseHeaderAndFoot(doc, map);
        // 第三步解析所有表格
        XWPFTable           table;
        Iterator<XWPFTable> itTable = doc.getTablesIterator();
        while (itTable.hasNext()) {
            table = itTable.next();
            if (table.getText().indexOf(START_STR) != -1) {
                parseThisTable(table, map);
            }
        }

    }

    /**
     * 解析页眉和页脚
     *
     * @param doc
     * @param map
     * @throws Exception
     */
    private void parseHeaderAndFoot(MyXWPFDocument doc, Map<String, Object> map) throws Exception {
        List<XWPFHeader> headerList = doc.getHeaderList();
        for (XWPFHeader xwpfHeader : headerList) {
            for (int i = 0; i < xwpfHeader.getListParagraph().size(); i++) {
                parseThisParagraph(xwpfHeader.getListParagraph().get(i), map);
            }
        }
        List<XWPFFooter> footerList = doc.getFooterList();
        for (XWPFFooter xwpfFooter : footerList) {
            for (int i = 0; i < xwpfFooter.getListParagraph().size(); i++) {
                parseThisParagraph(xwpfFooter.getListParagraph().get(i), map);
            }
        }

    }

}
