package org.jeecgframework.poi.excel.export.template;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import static org.jeecgframework.poi.util.PoiElUtil.*;

/**
 * 针对模板统计问题做统一处理
 * 1.处理模板之前统计需要SUM的数据以及位置
 * 2.遍历时统计数据
 * 3.遍历后设置数据
 * @author JueYue
 * @date 2016年6月19日
 */
public class TemplateSumHanlder {

    private List<TemplateSumEntity> sumList = new ArrayList<TemplateSumEntity>();

    /**
     * 统计计算所有的统计单元格
     */
    public void getAllSumCell(Sheet sheet) {
        Row row = null;
        int index = 0;
        while (index <= sheet.getLastRowNum()) {
            row = sheet.getRow(index++);
            if (row == null) {
                continue;
            }
            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                if (row.getCell(i) != null && row.getCell(i).getStringCellValue().contains(SUM)) {
                    addSumCellToList(row.getCell(i));
                }
            }
        }
    }

    private void addSumCellToList(Cell cell) {
        String cellValue = cell.getStringCellValue();
        int index = 0;
        while ((index = indexOfIgnoreCase(cellValue, SUM, index)) != -1) {
            TemplateSumEntity entity = new TemplateSumEntity();
            entity.setCellValue(cellValue);
            entity.setSumKey(getSumKey(cellValue, index++));
            entity.setCol(cell.getColumnIndex());
            entity.setRow(cell.getRowIndex());
            sumList.add(entity);
        }
    }

    /**
     * SUM:(key)
     * 
     * @param cellValue
     * @param index 
     * @return
     */
    private String getSumKey(String cellValue, int index) {
        return cellValue.substring(index + 4, cellValue.indexOf(")", index));
    }
    
    public void addListSizeToSumEntity(){
        
    }
    
    /**
     * 
     * @param rowIndex
     */
    public void findForeachList(int rowIndex){
        
    }

    private static int indexOfIgnoreCase(String str, String searchStr, int startPos) {
        if (str == null || searchStr == null) {
            return -1;
        }
        if (startPos < 0) {
            startPos = 0;
        }
        int endLimit = (str.length() - searchStr.length()) + 1;
        if (startPos > endLimit) {
            return -1;
        }
        if (searchStr.length() == 0) {
            return startPos;
        }
        for (int i = startPos; i < endLimit; i++) {
            if (str.regionMatches(true, i, searchStr, 0, searchStr.length())) {
                return i;
            }
        }
        return -1;
    }

    /**
     * 统计对象
     * @author JueYue
     * @date 2016年6月19日
     */
    protected class TemplateSumEntity {
        /**
         * CELL的值
         */
        private String cellValue;
        /**
         * 需要计算的KEY
         */
        private String sumKey;
        /**
         * 列
         */
        private int    col;
        /**
         * 行
         */
        private int    row;
        /**
         * 最后值
         */
        private Object value;

        public String getCellValue() {
            return cellValue;
        }

        public void setCellValue(String cellValue) {
            this.cellValue = cellValue;
        }

        public String getSumKey() {
            return sumKey;
        }

        public void setSumKey(String sumKey) {
            this.sumKey = sumKey;
        }

        public int getCol() {
            return col;
        }

        public void setCol(int col) {
            this.col = col;
        }

        public int getRow() {
            return row;
        }

        public void setRow(int row) {
            this.row = row;
        }

        public Object getValue() {
            return value;
        }

        public void setValue(Object value) {
            this.value = value;
        }

    }

}
