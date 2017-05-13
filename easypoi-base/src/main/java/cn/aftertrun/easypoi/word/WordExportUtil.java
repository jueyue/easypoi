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
package cn.aftertrun.easypoi.word;

import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import cn.aftertrun.easypoi.word.parse.ParseWord07;

/**
 * Word使用模板导出工具类
 * 
 * @author JueYue
 *  2013-11-16
 * @version 1.0
 */
public class WordExportUtil {

    private WordExportUtil() {

    }

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
     * @param document
     *            模板
     * @param map
     *            解析数据源
     */
    public static void exportWord07(XWPFDocument document,
                                    Map<String, Object> map) throws Exception {
        new ParseWord07().parseWord(document, map);
    }

}
