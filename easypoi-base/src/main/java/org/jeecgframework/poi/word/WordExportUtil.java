package org.jeecgframework.poi.word;

import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jeecgframework.poi.word.parse.ParseWord07;

/**
 * Word使用模板导出工具类
 * 
 * @author JueYue
 * @date 2013-11-16
 * @version 1.0
 */
public class WordExportUtil {

    /**
     * 解析Word2007版本
     * 
     * @param url
     *            模板地址
     * @param map
     *            解析数据源
     * @return
     */
    public static XWPFDocument exportWord07(String url, Map<String, Object> map) throws Exception {
        return new ParseWord07().parseWord(url, map);
    }

    /**
     * 解析Word2007版本
     * 
     * @param XWPFDocument
     *            模板
     * @param map
     *            解析数据源
     * @return
     */
    public static void exportWord07(XWPFDocument document, Map<String, Object> map)
                                                                                   throws Exception {
        new ParseWord07().parseWord(document, map);
    }

}
