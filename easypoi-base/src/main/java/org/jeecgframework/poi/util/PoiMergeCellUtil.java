package org.jeecgframework.poi.util;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jeecgframework.poi.excel.entity.params.MergeEntity;
import org.jeecgframework.poi.exception.excel.ExcelExportException;

/**
 * 纵向合并单元格工具类
 * @author JueYue
 *  2015年6月21日 上午11:21:40
 */
public final class PoiMergeCellUtil {

    private PoiMergeCellUtil() {
    }

    /**
     * 纵向合并相同内容的单元格
     * 
     * @param sheet
     * @param startRow 开始行
     * @param columns 需要处理的列
     */
    public static void mergeCells(Sheet sheet, int startRow, Integer... columns) {
        if (columns == null) {
            throw new ExcelExportException("至少需要处理1列");
        }
        Map<Integer, int[]> mergeMap = new HashMap<Integer, int[]>();
        for (int i = 0; i < columns.length; i++) {
            mergeMap.put(columns[i], null);
        }
        mergeCells(sheet, mergeMap, startRow, sheet.getLastRowNum());
    }

    /**
     * 纵向合并相同内容的单元格
     * 
     * @param sheet
     * @param mergeMap key--列,value--依赖的列,没有传空
     * @param startRow 开始行
     */
    public static void mergeCells(Sheet sheet, Map<Integer, int[]> mergeMap, int startRow) {
        mergeCells(sheet, mergeMap, startRow, sheet.getLastRowNum());
    }

    /**
     * 纵向合并相同内容的单元格
     * 
     * @param sheet
     * @param mergeMap key--列,value--依赖的列,没有传空
     * @param startRow 开始行
     * @param endRow 结束行
     */
    public static void mergeCells(Sheet sheet, Map<Integer, int[]> mergeMap, int startRow,
                                  int endRow) {
        Map<Integer, MergeEntity> mergeDataMap = new HashMap<Integer, MergeEntity>();
        if (mergeMap.size() == 0) {
            return;
        }
        Row row;
        Set<Integer> sets = mergeMap.keySet();
        String text;
        for (int i = startRow; i <= endRow; i++) {
            row = sheet.getRow(i);
            for (Integer index : sets) {
                if (row == null || row.getCell(index) == null) {
                    if (mergeDataMap.get(index) == null) {
                        continue;
                    }
                    if (mergeDataMap.get(index).getEndRow() == 0) {
                        mergeDataMap.get(index).setEndRow(i - 1);
                    }
                } else {
                    text = row.getCell(index).getStringCellValue();
                    if (StringUtils.isNotEmpty(text)) {
                        hanlderMergeCells(index, i, text, mergeDataMap, sheet, row.getCell(index),
                            mergeMap.get(index));
                    } else {
                        mergeCellOrContinue(index, mergeDataMap, sheet);
                    }
                }
            }
        }
        if (mergeDataMap.size() > 0) {
            for (Integer index : mergeDataMap.keySet()) {
                sheet.addMergedRegion(new CellRangeAddress(mergeDataMap.get(index).getStartRow(),
                    mergeDataMap.get(index).getEndRow(), index, index));
            }
        }

    }

    /**
     * 处理合并单元格
     * 
     * @param index
     * @param rowNum
     * @param text
     * @param mergeDataMap
     * @param sheet
     * @param cell
     * @param delys
     */
    private static void hanlderMergeCells(Integer index, int rowNum, String text,
                                          Map<Integer, MergeEntity> mergeDataMap, Sheet sheet,
                                          Cell cell, int[] delys) {
        if (mergeDataMap.containsKey(index)) {
            if (checkIsEqualByCellContents(mergeDataMap.get(index), text, cell, delys, rowNum)) {
                mergeDataMap.get(index).setEndRow(rowNum);
            } else {
                sheet.addMergedRegion(new CellRangeAddress(mergeDataMap.get(index).getStartRow(),
                    mergeDataMap.get(index).getEndRow(), index, index));
                mergeDataMap.put(index, createMergeEntity(text, rowNum, cell, delys));
            }
        } else {
            mergeDataMap.put(index, createMergeEntity(text, rowNum, cell, delys));
        }
    }

    /**
     * 字符为空的情况下判断
     * 
     * @param index
     * @param mergeDataMap
     * @param sheet
     */
    private static void mergeCellOrContinue(Integer index, Map<Integer, MergeEntity> mergeDataMap,
                                            Sheet sheet) {
        if (mergeDataMap.containsKey(index)
            && mergeDataMap.get(index).getEndRow() != mergeDataMap.get(index).getStartRow()) {
            sheet.addMergedRegion(new CellRangeAddress(mergeDataMap.get(index).getStartRow(),
                mergeDataMap.get(index).getEndRow(), index, index));
            mergeDataMap.remove(index);
        }
    }

    private static MergeEntity createMergeEntity(String text, int rowNum, Cell cell, int[] delys) {
        MergeEntity mergeEntity = new MergeEntity(text, rowNum, rowNum);
        // 存在依赖关系
        if (delys != null && delys.length != 0) {
            List<String> list = new ArrayList<String>(delys.length);
            mergeEntity.setRelyList(list);
            for (int i = 0; i < delys.length; i++) {
                list.add(getCellNotNullText(cell, delys[i], rowNum));
            }
        }
        return mergeEntity;
    }

    private static boolean checkIsEqualByCellContents(MergeEntity mergeEntity, String text,
                                                      Cell cell, int[] delys, int rowNum) {
        // 没有依赖关系
        if (delys == null || delys.length == 0) {
            return mergeEntity.getText().equals(text);
        }
        // 存在依赖关系 测试
        if (mergeEntity.getText().equals(text)) {
            for (int i = 0; i < delys.length; i++) {
                if (mergeEntity.getRelyList().get(i) == null || !mergeEntity.getRelyList().get(i)
                    .equals(getCellNotNullText(cell, delys[i], rowNum))) {
                    return false;
                }
            }
            return true;
        }
        return false;
    }

    /**
     * 获取一个单元格的值,确保这个单元格必须有值,不然向上查询
     * 
     * @param cell
     * @param index
     * @param rowNum
     * @return
     */
    private static String getCellNotNullText(Cell cell, int index, int rowNum) {
        if (cell == null || cell.getRow() == null) {
            return null;
        }
        if (cell.getRow().getCell(index) != null
            && StringUtils.isNotEmpty(cell.getRow().getCell(index).getStringCellValue())) {
            return cell.getRow().getCell(index).getStringCellValue();
        }
        return getCellNotNullText(cell.getRow().getSheet().getRow(--rowNum).getCell(index), index,
            rowNum);
    }

}
