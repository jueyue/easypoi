package org.jeecgframework.poi.excel;

import static org.junit.Assert.*;

import java.io.File;
import java.util.Collection;
import java.util.List;

import org.jeecgframework.poi.entity.CourseEntity;
import org.jeecgframework.poi.entity.StudentEntity;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.junit.Test;

public class ExcelImportUtilTest {

	@Test
	public void test() {
		ImportParams params = new ImportParams();
		params.setTitleRows(2);
		params.setHeadRows(2);
		List<CourseEntity> list = ExcelImportUtil.importExcel(new File(
				"d:/tt.xls"), CourseEntity.class, params);
		for (int i = 0; i < list.size(); i++) {
			System.out.println(list.get(i).getName());
		}
	}

}
