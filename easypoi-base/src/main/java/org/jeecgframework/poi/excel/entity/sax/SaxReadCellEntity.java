package org.jeecgframework.poi.excel.entity.sax;

import org.jeecgframework.poi.excel.entity.enmus.CellValueType;

/**
 * Cell 对象
 * @author JueYue
 * @date 2014年12月29日 下午10:12:57
 */
public class SaxReadCellEntity {
    /**
     * 值类型
     */
    private CellValueType cellType;
    /**
     * 值
     */
    private Object        value;

    public SaxReadCellEntity(CellValueType cellType, Object value) {
        this.cellType = cellType;
        this.value = value;
    }

    public CellValueType getCellType() {
        return cellType;
    }

    public void setCellType(CellValueType cellType) {
        this.cellType = cellType;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    @Override
    public String toString() {
        return "[type=" + cellType.toString() + ",value=" + value + "]";
    }

}
