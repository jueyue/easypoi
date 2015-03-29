package org.jeecgframework.poi.test.excel.test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.entity.vo.PoiBaseConstants;
import org.junit.Test;

public class ExcelExportForMap {
    /**
     * Map 测试
     */
    // @Test
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

            Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("测试", "测试"), entity,
                list);
            FileOutputStream fos = new FileOutputStream("d:/tt.xls");
            workbook.write(fos);
            fos.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testMany() {
        try {
            List<ExcelExportEntity> entity = new ArrayList<ExcelExportEntity>();
            for (int i = 0; i < 500; i++) {
                entity.add(new ExcelExportEntity("姓名" + i, "name" + i));
            }

            List<Map<String, String>> list = new ArrayList<Map<String, String>>();
            Map<String, String> map;
            for (int i = 0; i < 10; i++) {
                map = new HashMap<String, String>();
                for (int j = 0; j < 500; j++) {
                    map.put("name" + j, j + "_" + i);
                }
                list.add(map);
            }
            ExportParams params = new ExportParams("测试", "测试", ExcelType.XSSF);
            params.setFreezeCol(5);
            Workbook workbook = ExcelExportUtil.exportExcel(params, entity, list);
            FileOutputStream fos = new FileOutputStream("d:/tt.xlsx");
            workbook.write(fos);
            fos.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
