package org.jeecgframework.poi.test.excel.test;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelStyleType;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.excel.entity.vo.PoiBaseConstants;
import org.jeecgframework.poi.test.entity.statistics.StatisticEntity;
import org.jeecgframework.poi.test.excel.styler.ExcelExportStatisticStyler;
import org.junit.Test;

public class ExcelExportStatisticTest {
    @Test
    public void test() throws Exception {

        List<StatisticEntity> list = new ArrayList<StatisticEntity>();
        for (int i = 0; i < 20; i++) {
            StatisticEntity client = new StatisticEntity();
            client.setName("index" + i);
            client.setIntTest(1 + i);
            client.setLongTest(1 + i);
            client.setDoubleTest(1.2D + i);
            client.setBigDecimalTest(new BigDecimal(1.2 + i));
            client.setStringTest("12" + i);
            list.add(client);
        }
        Date start = new Date();
        ExportParams params = new ExportParams("2412312", "测试", ExcelType.XSSF);
        params.setStyle(ExcelExportStatisticStyler.class);
        Workbook workbook = ExcelExportUtil.exportExcel(params, StatisticEntity.class, list);
        System.out.println(new Date().getTime() - start.getTime());
        File savefile = new File("d:/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("d:/tt.xlsx");
        workbook.write(fos);
        fos.close();
    }
}
