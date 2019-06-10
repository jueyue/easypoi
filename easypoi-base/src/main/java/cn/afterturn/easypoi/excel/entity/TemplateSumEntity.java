package cn.afterturn.easypoi.excel.entity;

/**
 * 统计对象
 *
 * @author JueYue
 */
public class TemplateSumEntity {

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
    private double value;

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

    public double getValue() {
        return value;
    }

    public void setValue(double value) {
        this.value = value;
    }


}
