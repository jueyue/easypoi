package org.jeecgframework.poi.excel;

import java.io.File;
import java.util.Date;
import java.util.List;

import org.jeecgframework.poi.entity.CourseEntity;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.junit.Test;

public class ExcelImportUtilTest {

    @Test
    public void test() {
        ImportParams params = new ImportParams();
        params.setTitleRows(2);
        params.setHeadRows(2);
        //params.setSheetNum(9);
        params.setNeedSave(true);
        long start = new Date().getTime();
        List<CourseEntity> list = ExcelImportUtil.importExcel(new File("d:/tt.xls"),
            CourseEntity.class, params);
        System.out.println(list.size() + "-----" + (new Date().getTime() - start));
    }

}
