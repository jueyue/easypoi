/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 *   
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.jeecgframework.poi.util;

import java.util.Map;

import org.jeecgframework.poi.exception.excel.ExcelExportException;

/**
 * EasyPoi的el 表达式支持工具类
 * @author JueYue
 *  2015年4月25日 下午12:13:21
 */
public final class PoiElUtil {

	public static final String	LENGTH				= "le:";
	public static final String	FOREACH				= "fe:";
	public static final String	FOREACH_NOT_CREATE	= "!fe:";
	public static final String	FOREACH_AND_SHIFT	= "$fe:";
	public static final String	FOREACH_COL			= "#fe:";
	public static final String	FOREACH_COL_VALUE	= "v_fe:";
	public static final String	START_STR			= "{{";
	public static final String	END_STR				= "}}";
	public static final String	WRAP				= "]]";
	public static final String	NUMBER_SYMBOL		= "n:";
	public static final String	FORMAT_DATE			= "fd:";
	public static final String	FORMAT_NUMBER		= "fn:";
	public static final String	SUM		            = "sum:";
	public static final String	IF_DELETE			= "!if:";
	public static final String	EMPTY				= "";
	public static final String	CONST				= "'";
	public static final String	NULL				= "&NULL&";
	public static final String	LEFT_BRACKET		= "(";
	public static final String	RIGHT_BRACKET		= ")";

	private PoiElUtil() {
	}

	/**
	 * 解析字符串,支持 le,fd,fn,!if,三目
	 * @param text
	 * @param map
	 * @return
	 * @throws Exception 
	 */
	public static Object eval(String text, Map<String, Object> map) throws Exception {
		String tempText = new String(text);
		Object obj = innerEval(text, map);
		//如果没有被处理而且这个值找map中存在就处理这个值
		if (tempText.equals(obj.toString()) && map.containsKey(tempText.split("\\.")[0])) {
			return PoiPublicUtil.getParamsValue(tempText, map);
		}
		return obj;
	}

	/**
	 * 解析字符串,支持 le,fd,fn,!if,三目
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
		if (text.startsWith("'")) {
			return text.replace("'", "");
		}
		return text;
	}

	/**
	 * 是不是删除列
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
			Object first = eval(keys[0], map);
			Object second = eval(keys[2], map);
			return PoiFunctionUtil.isTrue(first, keys[1], second);
		}
		throw new ExcelExportException("判断参数不对");
	}

	/**
	 * 判断是不是常量
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
		int index = text.indexOf(prefix) + prefix.length();
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
	 * @return
	 * @throws Exception 
	 */
	private static Object trinocular(String text, Map<String, Object> map) throws Exception {
		//把多个空格变成一个空格
		text = text.replaceAll("\\s{1,}", " ").trim();
		String testText = text.substring(0, text.indexOf("?"));
		text = text.substring(text.indexOf("?") + 1, text.length()).trim();
		text = innerEval(text, map).toString();
		String[] keys = text.split(":");
		Object first = null, second = null;
		if(keys.length > 2){
		    if(keys[0].trim().contains("?")){
		        String trinocular = keys[0];
		        for (int i = 1; i < keys.length-1; i++) {
		            trinocular +=  ":" + keys[i];
                }
		        first = eval(trinocular, map);
		        second = eval(keys[keys.length-1].trim(), map);
		    } else {
		        first = eval(keys[0].trim(), map);
		        String trinocular = keys[1];
                for (int i = 2; i < keys.length; i++) {
                    trinocular +=  ":" + keys[i];
                }
		        second = eval(trinocular, map); 
		    }
		}else{
		    first = eval(keys[0].trim(), map);
	        second = eval(keys[1].trim(), map);
		}
		return isTrue(testText.split(" "), map) ? first : second;
	}

}
