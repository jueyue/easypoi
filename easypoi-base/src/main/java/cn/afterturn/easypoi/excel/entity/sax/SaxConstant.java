package cn.afterturn.easypoi.excel.entity.sax;

/**
 * @author by jueyue on 19-6-20.
 */
public interface SaxConstant {

    /**
     * Row 表达式
     */
    public static String ROW = "row";
    /**
     * cell 的位置
     */
    public static String ROW_COL = "r";
    /**
     * cell 表达式
     */
    public static String COL = "c";
    /**
     * tElement 表达式
     */
    public static String T_ELEMENT = "t";
    /**
     * cell 类型
     */
    public static String TYPE = "t";
    /**
     * style 缩写
     */
    public static String STYLE = "s";
    /**
     * String 缩写
     */
    public static String STRING = "s";
    /**
     * date 缩写
     */
    public static String DATE = "d";
    /**
     * number 缩写
     */
    public static String NUMBER = "n";
    /**
     * 计算表达式
     */
    public static String FORMULA = "str";
    /**
     * Boolean 缩写
     */
    public static String BOOLEAN = "b";
    /**
     * 类型值为“inlineStr”,表示这个单元格的字符串并没有用共享字符串池子的值
     */
    public static String INLINE_STR = "inlineStr";


    /**
     * value 缩写
     */
    public static String VALUE = "v";

}
