package org.jeecgframework.poi.excel.test;

import static org.junit.Assert.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jeecgframework.poi.entity.TeacherEntity;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;

public class ExcelExportForMap {

	@Test
	public void test() {
		try {
			List<ExcelExportEntity> entity = new ArrayList<ExcelExportEntity>();
			entity.add(new ExcelExportEntity("姓名", "name"));
			entity.add(new ExcelExportEntity("性别", "sex"));

			List<Map<String, String>> list = new ArrayList<Map<String, String>>();
			Map<String, String> map;
			for (int i = 0; i < 10; i++) {
				map = new HashMap<String, String>();
				map.put("name", "1" + i);
				map.put("sex", "2" + i);
				list.add(map);
			}

			HSSFWorkbook workbook = ExcelExportUtil.exportExcel(new ExportParams(
					"测试", "测试"), entity, list);
			FileOutputStream fos = new FileOutputStream("d:/tt.xls");
			workbook.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
