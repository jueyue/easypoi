package org.jeecgframework.poi.excel.test;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.entity.CourseEntity;
import org.jeecgframework.poi.entity.StudentEntity;
import org.jeecgframework.poi.entity.TeacherEntity;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.junit.Before;
import org.junit.Test;

/**
 * 关系导出测试 Created by JueYue on 14-4-19.
 */
public class ExcelExportUtilIdTest {

    List<CourseEntity>  list   = new ArrayList<CourseEntity>();
    List<TeacherEntity> telist = new ArrayList<TeacherEntity>();
    CourseEntity        courseEntity;

    @Before
    public void testBefor2() {
        TeacherEntity teacherEntity = new TeacherEntity();
        teacherEntity.setId("12132131231231231");
        teacherEntity.setName("你们");

        for (int i = 0; i < 3; i++) {
            telist.add(teacherEntity);
        }
    }

    @Before
    public void testBefore() {
        courseEntity = new CourseEntity();
        courseEntity.setId("1131");
        courseEntity.setName("小白");

        TeacherEntity teacherEntity = new TeacherEntity();
        teacherEntity.setId("12131231");
        teacherEntity.setName("你们");
        courseEntity.setTeacher(teacherEntity);

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
     * 单行表头导出测试
     * 
     * @throws Exception
     */
    @Test
    public void testExportExcel() throws Exception {
        ExportParams params = new ExportParams("2412312", "测试", "测试");
        params.setAddIndex(true);
        Workbook workbook = ExcelExportUtil.exportExcel(params, TeacherEntity.class, telist);
        FileOutputStream fos = new FileOutputStream("d:/tt.xls");
        workbook.write(fos);
        fos.close();
    }

}
