package org.jeecgframework.poi.test.excel.test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.test.entity.CourseEntity;
import org.jeecgframework.poi.test.entity.StudentEntity;
import org.jeecgframework.poi.test.entity.TeacherEntity;
import org.junit.Before;
import org.junit.Test;

/**
 * Created by jue on 14-4-19.
 */
public class ExcelExportUtilTest {

    List<CourseEntity> list = new ArrayList<CourseEntity>();
    CourseEntity       courseEntity;

    /**
     * 25W行导出测试
     * 
     * @throws Exception
     */
    @Test
    public void oneHundredThousandRowTest() throws Exception {

        ExportParams params = new ExportParams("2412312", "测试", ExcelType.XSSF);
        Date start = new Date();
        Workbook workbook = ExcelExportUtil.exportExcel(params, CourseEntity.class, list);
        System.out.println(new Date().getTime() - start.getTime());
        File savefile = new File("d:/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("d:/tt.xlsx");
        workbook.write(fos);
        fos.close();
//        savefile = new File("d:/1");
//        if (!savefile.exists()) {
//            savefile.setWritable(true, false);
//            savefile.mkdirs();
//        }
//        fos = new FileOutputStream("d:/1/tt3.xlsx");
//        workbook.write(fos);
//        fos.close();
    }

    @Before
    public void testBefore() {

        for (int i = 0; i < 20000; i++) {
            courseEntity = new CourseEntity();
            courseEntity.setId("1131");
            courseEntity.setName("海贼王必修(" + (i + 1) + ")");

            TeacherEntity teacherEntity = new TeacherEntity();
            teacherEntity.setId("12131231");
            teacherEntity.setName("路飞");
            courseEntity.setTeacher(teacherEntity);

            teacherEntity = new TeacherEntity();
            teacherEntity.setId("121312314312421131");
            teacherEntity.setName("老王");
            courseEntity.setShuxueteacher(teacherEntity);

            StudentEntity studentEntity = new StudentEntity();
            studentEntity.setId("11231");
            studentEntity.setName("撒旦法司法局");
            studentEntity.setBirthday(new Date());
            studentEntity.setSex(1);
            List<StudentEntity> studentList = new ArrayList<StudentEntity>();
            studentList.add(studentEntity);
            studentList.add(studentEntity);
            courseEntity.setStudents(studentList);
            list.add(courseEntity);
        }
    }

    /**
     * 基本导出测试
     * 
     * @throws Exception
     */
    //@Test
    public void testExportExcel() throws Exception {
        Date start = new Date();
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("2412312", "测试", "测试"),
            CourseEntity.class, list);
        System.out.println(new Date().getTime() - start.getTime());
        File savefile = new File("d:/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("d:/tt.xls");
        workbook.write(fos);
        fos.close();
    }

    /**
     * 单行表头测试
     * 
     * @throws Exception
     */
    //@Test
    public void testExportTitleExcel() throws Exception {
        Date start = new Date();
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("2412312", "测试"),
            CourseEntity.class, list);
        System.out.println(new Date().getTime() - start.getTime());
        File savefile = new File("d:/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("d:/tt.xls");
        workbook.write(fos);
        fos.close();
    }

    /**
     * 模板导出测试
     * 
     * @throws Exception
     */
    //@Test
    public void testTempExportExcel() throws Exception {
        TemplateExportParams params = new TemplateExportParams();
        params.setHeadingRows(2);
        params.setHeadingStartRow(2);
        params.setTemplateUrl("tt.xls");
        Workbook book = ExcelExportUtil.exportExcel(params, CourseEntity.class, list,
            new HashMap<String, Object>());
        File savefile = new File("d:/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("d:/t_tt.xls");
        book.write(fos);
        fos.close();
    }

}
