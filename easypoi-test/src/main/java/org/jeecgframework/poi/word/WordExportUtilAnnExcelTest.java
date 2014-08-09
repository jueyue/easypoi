package org.jeecgframework.poi.word;

import static org.junit.Assert.*;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jeecgframework.poi.entity.CourseEntity;
import org.jeecgframework.poi.entity.StudentEntity;
import org.jeecgframework.poi.entity.TeacherEntity;
import org.jeecgframework.poi.word.entity.Person;
import org.jeecgframework.poi.word.entity.WordImageEntity;
import org.jeecgframework.poi.word.entity.params.ExcelListEntity;
import org.junit.Before;
import org.junit.Test;

/**
 * 注解导出测试
 * 
 * @author JueYue
 * @date 2014年8月9日 下午11:36:33
 */
public class WordExportUtilAnnExcelTest {

	private static SimpleDateFormat format = new SimpleDateFormat("yyyy年MM月dd");
	
	List<CourseEntity> list = new ArrayList<CourseEntity>();
	CourseEntity courseEntity;

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
	 * 简单导出没有图片和Excel
	 */
	//@Test
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
		map.put("pList", new ExcelListEntity(list, Person.class));
		try {
			XWPFDocument doc = WordExportUtil.exportWord07(
					"org/jeecgframework/poi/word/doc/SimpleExcel.docx", map);
			FileOutputStream fos = new FileOutputStream("d:/simpleExcel.docx");
			doc.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	@Test
	public void WordExport() {
		Map<String, Object> map = new HashMap<String, Object>();
		map.put("department", "Jeecg");
		map.put("person", "JueYue");
		map.put("auditPerson", "JueYue");
		map.put("time", format.format(new Date()));
		map.put("cs", new ExcelListEntity(list, CourseEntity.class));
		try {
			XWPFDocument doc = WordExportUtil.exportWord07(
					"org/jeecgframework/poi/word/doc/Excel.docx", map);
			FileOutputStream fos = new FileOutputStream("d:/Excel.docx");
			doc.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
