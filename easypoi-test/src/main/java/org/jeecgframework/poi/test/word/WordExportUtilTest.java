package org.jeecgframework.poi.test.word;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jeecgframework.poi.word.WordExportUtil;
import org.jeecgframework.poi.word.entity.WordImageEntity;
import org.junit.Test;

public class WordExportUtilTest {

    private static SimpleDateFormat format = new SimpleDateFormat("yyyy年MM月dd");

    /**
     * 简单导出包含图片
     */
    //@Test
    public void imageWordExport() {
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("department", "Jeecg");
        map.put("person", "JueYue");
        map.put("time", format.format(new Date()));
        WordImageEntity image = new WordImageEntity();
        image.setHeight(200);
        image.setWidth(500);
        image.setUrl("org/jeecgframework/poi/word/img/testCode.png");
        image.setType(WordImageEntity.URL);
        map.put("testCode", image);
        try {
            XWPFDocument doc = WordExportUtil.exportWord07(
                "org/jeecgframework/poi/test/word/doc/Image.docx", map);
            FileOutputStream fos = new FileOutputStream("d:/image.docx");
            doc.write(fos);
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 简单导出没有图片和Excel
     */
    @Test
    public void SimpleWordExport() {
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("department", "Jeecg");
        map.put("person", "JueYue");
        map.put("time", format.format(new Date()));
        map.put("me","JueYue");
        map.put("date", "2015-01-03");
        try {
            XWPFDocument doc = WordExportUtil.exportWord07(
                "org/jeecgframework/poi/test/word/doc/Simple.docx", map);
            FileOutputStream fos = new FileOutputStream("d:/simple.docx");
            doc.write(fos);
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
