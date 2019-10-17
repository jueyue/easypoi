package cn.afterturn.easypoi.excel.entity.params;

import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;

import java.io.Serializable;
import java.util.Stack;

/**
 * 模板for each是的参数
 *
 * @author JueYue
 * 2015年4月29日 下午9:22:48
 */
@Data
public class ExcelForEachParams implements Serializable {

    /**
     *
     */
    private static final long          serialVersionUID = 1L;
    /**
     * key
     */
    private              String        name;
    /**
     * key
     */
    private              Stack<String> tempName;
    /**
     * 模板的cellStyle
     */
    private              CellStyle     cellStyle;
    /**
     * 行高
     */
    private              short         height;
    /**
     * 常量值
     */
    private              String        constValue;
    /**
     * 列合并
     */
    private              int           colspan          = 1;
    /**
     * 行合并
     */
    private              int           rowspan          = 1;
    /**
     * 行合并
     */
    private              boolean       collectCell;

    private boolean needSum;

    public ExcelForEachParams() {

    }

    public ExcelForEachParams(String name, CellStyle cellStyle, short height) {
        this.name = name;
        this.cellStyle = cellStyle;
        this.height = height;
    }

    public ExcelForEachParams(String name, CellStyle cellStyle, short height, boolean needSum) {
        this.name = name;
        this.cellStyle = cellStyle;
        this.height = height;
        this.needSum = needSum;
    }


}
