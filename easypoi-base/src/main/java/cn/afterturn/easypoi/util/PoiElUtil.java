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
package cn.afterturn.easypoi.util;

import cn.afterturn.easypoi.exception.excel.ExcelExportException;

import java.util.Arrays;
import java.util.Collections;
import java.util.Map;
import java.util.Stack;

/**
 * EasyPoi的el 表达式支持工具类
 *
 * @author JueYue
 * 2015年4月25日 下午12:13:21
 */
public final class PoiElUtil {

    public static final String LENGTH             = "le:";
    public static final String FOREACH            = "fe:";
    public static final String FOREACH_NOT_CREATE = "!fe:";
    public static final String FOREACH_AND_SHIFT  = "$fe:";
    public static final String FOREACH_COL        = "#fe:";
    public static final String FOREACH_COL_VALUE  = "v_fe:";
    public static final String START_STR          = "{{";
    public static final String END_STR            = "}}";
    public static final String WRAP               = "]]";
    public static final String NUMBER_SYMBOL      = "n:";
    public static final String STYLE_SELF         = "sy:";
    public static final String FORMAT_DATE        = "fd:";
    public static final String FORMAT_NUMBER      = "fn:";
    public static final String SUM                = "sum:";
    public static final String IF_DELETE          = "!if:";
    public static final String EMPTY              = "";
    public static final String CONST              = "'";
    public static final String NULL               = "&NULL&";
    public static final String LEFT_BRACKET       = "(";
    public static final String RIGHT_BRACKET      = ")";
    public static final String CAL                = "cal:";

    private PoiElUtil() {
    }

    /**
     * 解析字符串,支持 le,fd,fn,!if,三目
     *
     * @param text
     * @param map
     * @return
     * @throws Exception
     */
    public static Object eval(String text, Map<String, Object> map) throws Exception {
        String tempText = new String(text);
        Object obj      = innerEval(text, map);
        //如果没有被处理而且这个值找map中存在就处理这个值,找不到就返回空字符串
        if (tempText.equals(obj.toString())) {
            if (map.containsKey(tempText.split("\\.")[0])) {
                return PoiPublicUtil.getParamsValue(tempText, map);
            } else {
                return "";
            }
        }
        return obj;
    }

    /**
     * 解析字符串,支持 le,fd,fn,!if,三目  找不到返回原值
     *
     * @param text
     * @param map
     * @return
     * @throws Exception
     */
    public static Object evalNoParse(String text, Map<String, Object> map) throws Exception {
        String tempText = new String(text);
        Object obj      = innerEval(text, map);
        //如果没有被处理而且这个值找map中存在就处理这个值,找不到就返回空字符串
        if (tempText.equals(obj.toString())) {
            if (map.containsKey(tempText.split("\\.")[0])) {
                return PoiPublicUtil.getParamsValue(tempText, map);
            } else {
                return obj;
            }
        }
        return obj;
    }

    /**
     * 解析字符串,支持 le,fd,fn,!if,三目
     *
     * @param text
     * @param map
     * @return
     * @throws Exception
     */
    public static Object innerEval(String text, Map<String, Object> map) throws Exception {
        if (text.indexOf("?") != -1 && text.indexOf(":") != -1) {
            return trinocular(text, map);
        }
        if (text.indexOf(LENGTH) != -1) {
            return length(text, map);
        }
        if (text.indexOf(FORMAT_DATE) != -1) {
            return formatDate(text, map);
        }
        if (text.indexOf(FORMAT_NUMBER) != -1) {
            return formatNumber(text, map);
        }
        if (text.indexOf(IF_DELETE) != -1) {
            return ifDelete(text, map);
        }
        if (text.indexOf(CAL) != -1) {
            return calculate(text, map);
        }
        if (text.startsWith("'")) {
            return text.replace("'", "");
        }
        return text;
    }

    /**
     * 根据数据表达式计算结果
     *
     * @param text
     * @param map
     * @return
     */
    private static Object calculate(String text, Map<String, Object> map) throws Exception {
        //所有的数据都改成真是值
        //支持 + - * / () 所以按照这些字段切割
        text = text.replace(CAL, EMPTY);
        StringBuilder sb         = new StringBuilder();
        StringBuilder temp       = new StringBuilder();
        char[]        chars      = text.toCharArray();
        char[]        operations = new char[]{'+', '-', '*', '/', '(', ')', ' '};
        boolean       beforeMark = true;
        for (int i = 0; i < chars.length; i++) {
            if ( operations[0] == chars[i] || operations[1] == chars[i] || operations[2] == chars[i] || operations[3] == chars[i] || operations[4] == chars[i] || operations[5] == chars[i]) {
                if (temp.length() > 0) {
                    sb.append(evalNoParse(temp.toString().trim(), map).toString());
                    temp = new StringBuilder();
                }
                sb.append(chars[i]);
                beforeMark = true;
            } else if (beforeMark) {
                temp.append(chars[i]);
            }
        }
        if (temp.length() > 0) {
            sb.append(evalNoParse(temp.toString().trim(), map).toString());
        }
        return Calculator.conversion(sb.toString());
    }

    /**
     * 是不是删除列
     *
     * @param text
     * @param map
     * @return
     * @throws Exception
     */
    private static Object ifDelete(String text, Map<String, Object> map) throws Exception {
        //把多个空格变成一个空格
        text = text.replaceAll("\\s{1,}", " ").trim();
        String[] keys = getKey(IF_DELETE, text).split(" ");
        text = text.replace(IF_DELETE, EMPTY);
        return isTrue(keys, map);
    }

    /**
     * 是不是真
     *
     * @param keys
     * @param map
     * @return
     * @throws Exception
     */
    private static Boolean isTrue(String[] keys, Map<String, Object> map) throws Exception {
        if (keys.length == 1) {
            String constant = null;
            if ((constant = isConstant(keys[0])) != null) {
                return Boolean.valueOf(constant);
            }
            return Boolean.valueOf(PoiPublicUtil.getParamsValue(keys[0], map).toString());
        }
        if (keys.length == 3) {
            Object first  = evalNoParse(keys[0], map);
            Object second = evalNoParse(keys[2], map);
            return PoiFunctionUtil.isTrue(first, keys[1], second);
        }
        throw new ExcelExportException("判断参数不对");
    }

    /**
     * 判断是不是常量
     *
     * @param param
     * @return
     */
    private static String isConstant(String param) {
        if (param.indexOf("'") != -1) {
            return param.replace("'", "");
        }
        return null;
    }

    /**
     * 格式化数字
     *
     * @param text
     * @param map
     * @return
     * @throws Exception
     */
    private static Object formatNumber(String text, Map<String, Object> map) throws Exception {
        String[] key = getKey(FORMAT_NUMBER, text).split(";");
        text = text.replace(FORMAT_NUMBER, EMPTY);
        return innerEval(
                replacinnerEvalue(text,
                        PoiFunctionUtil.formatNumber(PoiPublicUtil.getParamsValue(key[0], map), key[1])),
                map);
    }

    /**
     * 格式化时间
     *
     * @param text
     * @param map
     * @return
     * @throws Exception
     */
    private static Object formatDate(String text, Map<String, Object> map) throws Exception {
        String[] key = getKey(FORMAT_DATE, text).split(";");
        text = text.replace(FORMAT_DATE, EMPTY);
        return innerEval(
                replacinnerEvalue(text,
                        PoiFunctionUtil.formatDate(PoiPublicUtil.getParamsValue(key[0], map), key[1])),
                map);
    }

    /**
     * 计算这个的长度
     *
     * @param text
     * @param map
     * @throws Exception
     */
    private static Object length(String text, Map<String, Object> map) throws Exception {
        String key = getKey(LENGTH, text);
        text = text.replace(LENGTH, EMPTY);
        Object val = PoiPublicUtil.getParamsValue(key, map);
        return innerEval(replacinnerEvalue(text, PoiFunctionUtil.length(val)), map);
    }

    private static String replacinnerEvalue(String text, Object val) {
        StringBuilder sb = new StringBuilder();
        sb.append(text.substring(0, text.indexOf(LEFT_BRACKET)));
        sb.append(" ");
        sb.append(val);
        sb.append(" ");
        sb.append(text.substring(text.indexOf(RIGHT_BRACKET) + 1, text.length()));
        return sb.toString().trim();
    }

    private static String getKey(String prefix, String text) {
        int leftBracket = 1, rigthBracket = 0, position = 0;
        int index       = text.indexOf(prefix) + prefix.length();
        while (text.charAt(index) == " ".charAt(0)) {
            text = text.substring(0, index) + text.substring(index + 1, text.length());
        }
        for (int i = text.indexOf(prefix + LEFT_BRACKET) + prefix.length() + 1; i < text
                .length(); i++) {
            if (text.charAt(i) == LEFT_BRACKET.charAt(0)) {
                leftBracket++;
            }
            if (text.charAt(i) == RIGHT_BRACKET.charAt(0)) {
                rigthBracket++;
            }
            if (leftBracket == rigthBracket) {
                position = i;
                break;
            }
        }
        return text.substring(text.indexOf(prefix + LEFT_BRACKET) + 1 + prefix.length(), position)
                .trim();
    }

    /**
     * 三目运算
     *
     * @return
     * @throws Exception
     */
    private static Object trinocular(String text, Map<String, Object> map) throws Exception {
        //把多个空格变成一个空格
        text = text.replaceAll("\\s{1,}", " ").trim();
        String testText = text.substring(0, text.indexOf("?"));
        text = text.substring(text.indexOf("?") + 1, text.length()).trim();
        text = innerEval(text, map).toString();
        String[] keys  = text.split(":");
        Object   first = null, second = null;
        if (keys.length > 2) {
            if (keys[0].trim().contains("?")) {
                String trinocular = keys[0];
                for (int i = 1; i < keys.length - 1; i++) {
                    trinocular += ":" + keys[i];
                }
                first = evalNoParse(trinocular, map);
                second = evalNoParse(keys[keys.length - 1].trim(), map);
            } else {
                first = evalNoParse(keys[0].trim(), map);
                String trinocular = keys[1];
                for (int i = 2; i < keys.length; i++) {
                    trinocular += ":" + keys[i];
                }
                second = evalNoParse(trinocular, map);
            }
        } else {
            first = evalNoParse(keys[0].trim(), map);
            second = evalNoParse(keys[1].trim(), map);
        }
        return isTrue(testText.split(" "), map) ? first : second;
    }


    /**
     * 算数表达式求值
     * 直接调用Calculator的类方法conversion()
     * 传入算数表达式，将返回一个浮点值结果
     * 如果计算过程错误，将返回一个NaN
     */
    private static class Calculator {
        private Stack<String>    postfixStack   = new Stack<String>();// 后缀式栈
        private Stack<Character> opStack        = new Stack<Character>();// 运算符栈
        private int[]            operatPriority = new int[]{0, 3, 2, 1, -1, 1, 0, 2};// 运用运算符ASCII码-40做索引的运算符优先级

        public static Object conversion(String expression) {
            double     result = 0;
            Calculator cal    = new Calculator();
            try {
                expression = transform(expression);
                result = cal.calculate(expression);
            } catch (Exception e) {
                // e.printStackTrace();
                // 运算错误返回NaN
                return "NaN";
            }
            return result;
        }

        /**
         * 将表达式中负数的符号更改
         *
         * @param expression 例如-2+-1*(-3E-2)-(-1) 被转为 ~2+~1*(~3E~2)-(~1)
         * @return
         */
        private static String transform(String expression) {
            char[] arr = expression.toCharArray();
            for (int i = 0; i < arr.length; i++) {
                if (arr[i] == '-') {
                    if (i == 0) {
                        arr[i] = '~';
                    } else {
                        char c = arr[i - 1];
                        if (c == '+' || c == '-' || c == '*' || c == '/' || c == '(' || c == 'E' || c == 'e') {
                            arr[i] = '~';
                        }
                    }
                }
            }
            if (arr[0] == '~' || arr[1] == '(') {
                arr[0] = '-';
                return "0" + new String(arr);
            } else {
                return new String(arr);
            }
        }

        /**
         * 按照给定的表达式计算
         *
         * @param expression 要计算的表达式例如:5+12*(3+5)/7
         * @return
         */
        public double calculate(String expression) {
            Stack<String> resultStack = new Stack<String>();
            prepare(expression);
            Collections.reverse(postfixStack);
            String firstValue, secondValue, currentValue;
            while (!postfixStack.isEmpty()) {
                currentValue = postfixStack.pop();
                if (!isOperator(currentValue.charAt(0))) {
                    currentValue = currentValue.replace("~", "-");
                    resultStack.push(currentValue);
                } else {// 如果是运算符则从操作数栈中取两个值和该数值一起参与运算
                    secondValue = resultStack.pop();
                    firstValue = resultStack.pop();

                    // 将负数标记符改为负号
                    firstValue = firstValue.replace("~", "-");
                    secondValue = secondValue.replace("~", "-");

                    String tempResult = calculate(firstValue, secondValue, currentValue.charAt(0));
                    resultStack.push(tempResult);
                }
            }
            return Double.valueOf(resultStack.pop());
        }

        /**
         * 数据准备阶段将表达式转换成为后缀式栈
         *
         * @param expression
         */
        private void prepare(String expression) {
            opStack.push(',');// 运算符放入栈底元素逗号，此符号优先级最低
            char[] arr          = expression.toCharArray();
            int    currentIndex = 0;// 当前字符的位置
            int    count        = 0;// 上次算术运算符到本次算术运算符的字符的长度便于或者之间的数值
            char   currentOp, peekOp;// 当前操作符和栈顶操作符
            for (int i = 0; i < arr.length; i++) {
                currentOp = arr[i];
                if (isOperator(currentOp)) {// 如果当前字符是运算符
                    if (count > 0) {
                        postfixStack.push(new String(arr, currentIndex, count));// 取两个运算符之间的数字
                    }
                    peekOp = opStack.peek();
                    if (currentOp == ')') {// 遇到反括号则将运算符栈中的元素移除到后缀式栈中直到遇到左括号
                        while (opStack.peek() != '(') {
                            postfixStack.push(String.valueOf(opStack.pop()));
                        }
                        opStack.pop();
                    } else {
                        while (currentOp != '(' && peekOp != ',' && compare(currentOp, peekOp)) {
                            postfixStack.push(String.valueOf(opStack.pop()));
                            peekOp = opStack.peek();
                        }
                        opStack.push(currentOp);
                    }
                    count = 0;
                    currentIndex = i + 1;
                } else {
                    count++;
                }
            }
            if (count > 1 || (count == 1 && !isOperator(arr[currentIndex]))) {// 最后一个字符不是括号或者其他运算符的则加入后缀式栈中
                postfixStack.push(new String(arr, currentIndex, count));
            }

            while (opStack.peek() != ',') {
                postfixStack.push(String.valueOf(opStack.pop()));// 将操作符栈中的剩余的元素添加到后缀式栈中
            }
        }

        /**
         * 判断是否为算术符号
         *
         * @param c
         * @return
         */
        private boolean isOperator(char c) {
            return c == '+' || c == '-' || c == '*' || c == '/' || c == '(' || c == ')';
        }

        /**
         * 利用ASCII码-40做下标去算术符号优先级
         *
         * @param cur
         * @param peek
         * @return
         */
        public boolean compare(char cur, char peek) {// 如果是peek优先级高于cur，返回true，默认都是peek优先级要低
            boolean result = false;
            if (operatPriority[(peek) - 40] >= operatPriority[(cur) - 40]) {
                result = true;
            }
            return result;
        }

        /**
         * 按照给定的算术运算符做计算
         *
         * @param firstValue
         * @param secondValue
         * @param currentOp
         * @return
         */
        private String calculate(String firstValue, String secondValue, char currentOp) {
            String result = "";
            switch (currentOp) {
                case '+':
                    result = String.valueOf(ArithHelper.add(firstValue, secondValue));
                    break;
                case '-':
                    result = String.valueOf(ArithHelper.sub(firstValue, secondValue));
                    break;
                case '*':
                    result = String.valueOf(ArithHelper.mul(firstValue, secondValue));
                    break;
                case '/':
                    result = String.valueOf(ArithHelper.div(firstValue, secondValue));
                    break;
            }
            return result;
        }
    }


    private static class ArithHelper {

        // 默认除法运算精度
        private static final int DEF_DIV_SCALE = 16;

        // 这个类不能实例化
        private ArithHelper() {
        }

        /**
         * 提供精确的加法运算。
         *
         * @param v1 被加数
         * @param v2 加数
         * @return 两个参数的和
         */

        public static double add(double v1, double v2) {
            java.math.BigDecimal b1 = new java.math.BigDecimal(Double.toString(v1));
            java.math.BigDecimal b2 = new java.math.BigDecimal(Double.toString(v2));
            return b1.add(b2).doubleValue();
        }

        public static double add(String v1, String v2) {
            java.math.BigDecimal b1 = new java.math.BigDecimal(v1);
            java.math.BigDecimal b2 = new java.math.BigDecimal(v2);
            return b1.add(b2).doubleValue();
        }

        /**
         * 提供精确的减法运算。
         *
         * @param v1 被减数
         * @param v2 减数
         * @return 两个参数的差
         */

        public static double sub(double v1, double v2) {
            java.math.BigDecimal b1 = new java.math.BigDecimal(Double.toString(v1));
            java.math.BigDecimal b2 = new java.math.BigDecimal(Double.toString(v2));
            return b1.subtract(b2).doubleValue();
        }

        public static double sub(String v1, String v2) {
            java.math.BigDecimal b1 = new java.math.BigDecimal(v1);
            java.math.BigDecimal b2 = new java.math.BigDecimal(v2);
            return b1.subtract(b2).doubleValue();
        }

        /**
         * 提供精确的乘法运算。
         *
         * @param v1 被乘数
         * @param v2 乘数
         * @return 两个参数的积
         */

        public static double mul(double v1, double v2) {
            java.math.BigDecimal b1 = new java.math.BigDecimal(Double.toString(v1));
            java.math.BigDecimal b2 = new java.math.BigDecimal(Double.toString(v2));
            return b1.multiply(b2).doubleValue();
        }

        public static double mul(String v1, String v2) {
            java.math.BigDecimal b1 = new java.math.BigDecimal(v1);
            java.math.BigDecimal b2 = new java.math.BigDecimal(v2);
            return b1.multiply(b2).doubleValue();
        }

        /**
         * 提供（相对）精确的除法运算，当发生除不尽的情况时，精确到 小数点以后10位，以后的数字四舍五入。
         *
         * @param v1 被除数
         * @param v2 除数
         * @return 两个参数的商
         */

        public static double div(double v1, double v2) {
            return div(v1, v2, DEF_DIV_SCALE);
        }

        public static double div(String v1, String v2) {
            java.math.BigDecimal b1 = new java.math.BigDecimal(v1);
            java.math.BigDecimal b2 = new java.math.BigDecimal(v2);
            return b1.divide(b2, DEF_DIV_SCALE, java.math.BigDecimal.ROUND_HALF_UP).doubleValue();
        }

        /**
         * 提供（相对）精确的除法运算。当发生除不尽的情况时，由scale参数指 定精度，以后的数字四舍五入。
         *
         * @param v1    被除数
         * @param v2    除数
         * @param scale 表示表示需要精确到小数点以后几位。
         * @return 两个参数的商
         */

        public static double div(double v1, double v2, int scale) {
            if (scale < 0) {
                throw new IllegalArgumentException("The   scale   must   be   a   positive   integer   or   zero");
            }
            java.math.BigDecimal b1 = new java.math.BigDecimal(Double.toString(v1));
            java.math.BigDecimal b2 = new java.math.BigDecimal(Double.toString(v2));
            return b1.divide(b2, scale, java.math.BigDecimal.ROUND_HALF_UP).doubleValue();
        }

        /**
         * 提供精确的小数位四舍五入处理。
         *
         * @param v     需要四舍五入的数字
         * @param scale 小数点后保留几位
         * @return 四舍五入后的结果
         */

        public static double round(double v, int scale) {
            if (scale < 0) {
                throw new IllegalArgumentException("The   scale   must   be   a   positive   integer   or   zero");
            }
            java.math.BigDecimal b   = new java.math.BigDecimal(Double.toString(v));
            java.math.BigDecimal one = new java.math.BigDecimal("1");
            return b.divide(one, scale, java.math.BigDecimal.ROUND_HALF_UP).doubleValue();
        }

        public static double round(String v, int scale) {
            if (scale < 0) {
                throw new IllegalArgumentException("The   scale   must   be   a   positive   integer   or   zero");
            }
            java.math.BigDecimal b   = new java.math.BigDecimal(v);
            java.math.BigDecimal one = new java.math.BigDecimal("1");
            return b.divide(one, scale, java.math.BigDecimal.ROUND_HALF_UP).doubleValue();
        }
    }
}
