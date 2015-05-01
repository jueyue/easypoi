package org.jeecgframework.poi.test.word;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jeecgframework.poi.test.word.entity.Person;
import org.jeecgframework.poi.word.WordExportUtil;
import org.jeecgframework.poi.word.entity.WordImageEntity;
import org.junit.Test;

public class WordExportUtilBaseExcelTest {

    private static SimpleDateFormat format = new SimpleDateFormat("yyyy年MM月dd");

    /**
     * 简单导出包含图片
     */
    @Test
    public void imageWordExport() {
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("department", "Jeecg");
        map.put("person", "JueYue");
        map.put("time", format.format(new Date()));
        WordImageEntity image = new WordImageEntity();
        image.setHeight(200);
        image.setWidth(500);
        image.setUrl("org/jeecgframework/poi/test/word/img/testCode.png");
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
        map.put("auditPerson", "JueYue");
        map.put("time", format.format(new Date()));
        List<Person> list = new ArrayList<Person>();
        Person p = new Person();
        p.setName("小明");
        p.setTel("18711111111");
        p.setEmail("18711111111@139.com");
        list.add(p);
        p = new Person();
        p.setName("小红");
        p.setTel("18711111112");
        p.setEmail("18711111112@139.com");
        list.add(p);
        map.put("pList", list);
        try {
            XWPFDocument doc = WordExportUtil.exportWord07(
                "org/jeecgframework/poi/test/word/doc/SimpleExcel.docx", map);
            FileOutputStream fos = new FileOutputStream("d:/simpleExcel.docx");
            doc.write(fos);
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
