package org.jeecgframework.poi.test.excel.template;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerColorImpl;
import org.jeecgframework.poi.test.entity.temp.BudgetAccountsEntity;
import org.jeecgframework.poi.test.entity.temp.PayeeEntity;
import org.jeecgframework.poi.test.entity.temp.TemplateExcelExportEntity;
import org.junit.Test;

import com.google.common.collect.Lists;

public class TemplateForEachTest {

    @Test
    public void test() throws Exception {
        TemplateExportParams params = new TemplateExportParams(
            "org/jeecgframework/poi/test/excel/doc/foreach.xlsx");
        Map<String, Object> map = new HashMap<String, Object>();
        List<TemplateExcelExportEntity> list = new ArrayList<TemplateExcelExportEntity>();

        for (int i = 0; i < 4; i++) {
            TemplateExcelExportEntity entity = new TemplateExcelExportEntity();
            entity.setIndex(i + 1 + "");
            entity.setAccountType("开源项目");
            entity.setProjectName("EasyPoi " + i + "期");
            entity.setAmountApplied(i * 10000 + "");
            entity.setApprovedAmount((i + 1) * 10000 - 100 + "");
            list.add(entity);
        }
        map.put("entitylist", list);
        List<Map<String, Object>> mapList = new ArrayList<Map<String, Object>>();
        for (int i = 0; i < 4; i++) {
            Map<String, Object> testMap = new HashMap<String, Object>();

            testMap.put("id", i);
            testMap.put("name", "小明" + i);
            testMap.put("sex", "man");
            mapList.add(testMap);
        }
        map.put("maplist", mapList);
        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        File savefile = new File("d:/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("d:/foreach.xlsx");
        workbook.write(fos);
        fos.close();
        workbook = ExcelExportUtil.exportExcel(params, map);
        fos = new FileOutputStream("d:/foreach2.xlsx");
        workbook.write(fos);
        fos.close();
    }

}
