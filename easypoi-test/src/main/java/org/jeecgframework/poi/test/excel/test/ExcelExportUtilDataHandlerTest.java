package org.jeecgframework.poi.test.excel.test;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.test.entity.CourseEntity;
import org.jeecgframework.poi.test.entity.StudentEntity;
import org.jeecgframework.poi.test.entity.TeacherEntity;
import org.junit.Before;
import org.junit.Test;

/**
 * 数据处理测试 Created by jue on 14-4-19.
 */
public class ExcelExportUtilDataHandlerTest {

    List<CourseEntity> list = new ArrayList<CourseEntity>();
    CourseEntity       courseEntity;

    @Before
    public void testBefore() {
        courseEntity = new CourseEntity();
        courseEntity.setId("1131");
        courseEntity.setName("小白");

        TeacherEntity teacherEntity = new TeacherEntity();
        teacherEntity.setId("12131231");
        teacherEntity.setName("你们");
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

        for (int i = 0; i < 3; i++) {
            list.add(courseEntity);
        }
    }

    /**
     * 基本导出测试
     * 
     * @throws Exception
     */
    @Test
    public void testExportExcel() throws Exception {
        ExportParams exportParams = new ExportParams("2412312", "测试", "测试");
        CourseHanlder hanlder = new CourseHanlder();
        hanlder.setNeedHandlerFields(new String[] { "课程名称" });
        exportParams.setDataHanlder(hanlder);
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, CourseEntity.class, list);
        FileOutputStream fos = new FileOutputStream("d:/tt.xls");
        workbook.write(fos);
        fos.close();
    }
}
