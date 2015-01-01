package org.jeecgframework.poi.excel.test;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.entity.CourseEntity;
import org.jeecgframework.poi.entity.MsgClient;
import org.jeecgframework.poi.entity.MsgClientGroup;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.excel.entity.vo.PoiBaseConstants;
import org.junit.Test;
/**
 * 当行大数据量测试
 * @author JueYue
 * @date 2014年12月13日 下午3:42:57
 */
public class ExcelExportMsgClient {

    @Test
    public void test() throws Exception {

        List<MsgClient> list = new ArrayList<MsgClient>();
        for (int i = 0; i < 50000; i++) {
            MsgClient client = new MsgClient();
            client.setBirthday(new Date());
            client.setClientName("小明" + i);
            client.setClientPhone("18797" + i);
            client.setCreateBy("jueyue");
            client.setId("1" + i);
            client.setRemark("测试" + i);
            MsgClientGroup group = new MsgClientGroup();
            group.setGroupName("测试" + i);
            client.setGroup(group);
            list.add(client);
        }
        Date start = new Date();
        ExportParams params = new ExportParams("2412312", "测试", ExcelType.XSSF);
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
