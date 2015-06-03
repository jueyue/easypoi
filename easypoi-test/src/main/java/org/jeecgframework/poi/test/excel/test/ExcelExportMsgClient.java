package org.jeecgframework.poi.test.excel.test;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.excel.entity.vo.PoiBaseConstants;
import org.jeecgframework.poi.test.entity.CourseEntity;
import org.jeecgframework.poi.test.entity.MsgClient;
import org.jeecgframework.poi.test.entity.MsgClientGroup;
import org.junit.Test;
/**
 * 当行大数据量测试
 * @author JueYue
 * @date 2014年12月13日 下午3:42:57
 */
public class ExcelExportMsgClient {

    @Test
    public void test() throws Exception {

        Field[] fields =  MsgClient.class.getFields();
        for (int i = 0; i < fields.length; i++) {
            Excel excel = fields[i].getAnnotation(Excel.class);
            System.out.println(excel);
        }
        
        List<MsgClient> list = new ArrayList<MsgClient>();
        for (int i = 0; i < 1000000; i++) {
            MsgClient client = new MsgClient();
            client.setBirthday(new Date());
            client.setClientName("小明" + i);
            client.setClientPhone("18797" + i);
            client.setCreateBy("JueYue");
            client.setId("1" + i);
            client.setRemark("测试" + i);
            MsgClientGroup group = new MsgClientGroup();
            group.setGroupName("测试" + i);
            client.setGroup(group);
            list.add(client);
        }
        Date start = new Date();
        ExportParams params = new ExportParams("2412312", "测试", ExcelType.XSSF);
        params.setFreezeCol(2);
        Workbook workbook = ExcelExportUtil.exportExcel(params, MsgClient.class, list);
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
