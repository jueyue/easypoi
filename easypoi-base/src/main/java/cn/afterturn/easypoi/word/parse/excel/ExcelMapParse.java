/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License. You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under the License
 * is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
 * or implied. See the License for the specific language governing permissions and limitations under
 * the License.
 */
package cn.afterturn.easypoi.word.parse.excel;

import cn.afterturn.easypoi.entity.ImageEntity;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import cn.afterturn.easypoi.word.entity.MyXWPFDocument;
import com.google.common.collect.Maps;

import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import cn.afterturn.easypoi.util.PoiWordStyleUtil;

import static cn.afterturn.easypoi.util.PoiElUtil.EMPTY;
import static cn.afterturn.easypoi.util.PoiElUtil.END_STR;
import static cn.afterturn.easypoi.util.PoiElUtil.FOREACH;
import static cn.afterturn.easypoi.util.PoiElUtil.FOREACH_AND_SHIFT;
import static cn.afterturn.easypoi.util.PoiElUtil.FOREACH_NOT_CREATE;
import static cn.afterturn.easypoi.util.PoiElUtil.START_STR;
import static cn.afterturn.easypoi.util.PoiElUtil.eval;

/**
 * 处理和生成Map 类型的数据变成表格
 * @author JueYue
 *  2014年8月9日 下午10:28:46
 */
public final class ExcelMapParse {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelMapParse.class);

    /**
     * 添加图片
     *
     * @param obj
     * @param currentRun
     * @throws Exception
     * @author JueYue
     * 2013-11-20
     */
    public static void addAnImage(ImageEntity obj, XWPFRun currentRun) {
        try {
            Object[] isAndType = PoiPublicUtil.getIsAndType(obj);
            String   picId;
            picId = currentRun.getDocument().addPictureData((byte[]) isAndType[0],
                    (Integer) isAndType[1]);
            ((MyXWPFDocument) currentRun.getDocument()).createPicture(currentRun,
                    picId, currentRun.getDocument()
                            .getNextPicNameNumber((Integer) isAndType[1]),
                    obj.getWidth(), obj.getHeight());

        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }

    }

    /**
     * 解析参数行,获取参数列表
     *
     * @author JueYue
     *  2013-11-18
     * @param currentRow
     * @return
     */
    private static String[] parseCurrentRowGetParams(XWPFTableRow currentRow) {
        List<XWPFTableCell> cells = currentRow.getTableCells();
        String[] params = new String[cells.size()];
        String text;
        for (int i = 0; i < cells.size(); i++) {
            text = cells.get(i).getText();
            params[i] = text == null ? ""
                    : text.trim().replace(START_STR, EMPTY).replace(END_STR, EMPTY);
        }
        return params;
    }

    /**
     * 解析下一行,并且生成更多的行
     * @param table
     * @param index
     * @param list
     */
    public static void parseNextRowAndAddRow(XWPFTable table, int index,
                                             List<Object> list) throws Exception {
        XWPFTableRow currentRow = table.getRow(index);
        String[] params = parseCurrentRowGetParams(currentRow);
        String listname = params[0];
        boolean isCreate = !listname.contains(FOREACH_NOT_CREATE);
        listname = listname.replace(FOREACH_NOT_CREATE, EMPTY).replace(FOREACH_AND_SHIFT, EMPTY)
                .replace(FOREACH, EMPTY).replace(START_STR, EMPTY);
        String[] keys = listname.replaceAll("\\s{1,}", " ").trim().split(" ");
        params[0] = keys[1];
        //保存这一行的样式是-后面好统一设置
        List<XWPFTableCell> tempCellList = new ArrayList<XWPFTableCell>();
        tempCellList.addAll(table.getRow(index).getTableCells());
        int cellIndex = 0;
        Map<String, Object> tempMap = Maps.newHashMap();
        LOGGER.debug("start for each data list :{}", list.size());
        for (Object obj : list) {
            currentRow = isCreate ? table.insertNewTableRow(index++) : table.getRow(index++);
            tempMap.put("t", obj);
            for (cellIndex = 0; cellIndex < currentRow.getTableCells().size(); cellIndex++) {
                String val = eval(params[cellIndex], tempMap).toString();
                currentRow.getTableCells().get(cellIndex).setText("");
                PoiWordStyleUtil.copyCellAndSetValue(tempCellList.get(cellIndex),
                        currentRow.getTableCells().get(cellIndex), val);
            }

            for (; cellIndex < params.length; cellIndex++) {
                String val = eval(params[cellIndex], tempMap).toString();
                PoiWordStyleUtil.copyCellAndSetValue(tempCellList.get(cellIndex),
                        currentRow.createCell(), val);
            }
        }
        table.removeRow(index);

    }

}
