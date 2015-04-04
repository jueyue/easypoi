package org.jeecgframework.poi.test.excel.template;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerBorderImpl;
import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerColorImpl;
import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerDefaultImpl;
import org.jeecgframework.poi.test.entity.statistics.StatisticEntity;
import org.jeecgframework.poi.test.entity.temp.BudgetAccountsEntity;
import org.jeecgframework.poi.test.entity.temp.PayeeEntity;
import org.jeecgframework.poi.test.entity.temp.TemplateExcelExportEntity;
import org.junit.Test;

import com.google.common.collect.Lists;

public class TemplateExcelExportTest {

    @Test
    public void test() throws Exception {
        TemplateExportParams params = new TemplateExportParams(
            "org/jeecgframework/poi/test/excel/doc/专项支出用款申请书.xls");
        params.setHeadingStartRow(3);
        params.setHeadingRows(2);
        params.setStyle(ExcelExportStylerColorImpl.class);
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("date", "2014-12-25");
        map.put("money", 2000000.00);
        map.put("upperMoney", "贰佰万");
        map.put("company", "执笔潜行科技有限公司");
        map.put("bureau", "财政局");
        map.put("person", "JueYue");
        map.put("phone", "1879740****");

        List<TemplateExcelExportEntity> list = new ArrayList<TemplateExcelExportEntity>();

        for (int i = 0; i < 4; i++) {
            TemplateExcelExportEntity entity = new TemplateExcelExportEntity();
            entity.setIndex(i + 1 + "");
            entity.setAccountType("开源项目");
            entity.setProjectName("EasyPoi " + i + "期");
            entity.setAmountApplied(i * 10000 + "");
            entity.setApprovedAmount((i + 1) * 10000 - 100 + "");
            List<BudgetAccountsEntity> budgetAccounts = Lists.newArrayList();
            for (int j = 0; j < 1; j++) {
                BudgetAccountsEntity accountsEntity = new BudgetAccountsEntity();
                accountsEntity.setCode("A001");
                accountsEntity.setName("设计");
                budgetAccounts.add(accountsEntity);
                accountsEntity = new BudgetAccountsEntity();
                accountsEntity.setCode("A002");
                accountsEntity.setName("开发");
                budgetAccounts.add(accountsEntity);
            }
            entity.setBudgetAccounts(budgetAccounts);
            PayeeEntity payeeEntity = new PayeeEntity();
            payeeEntity.setBankAccount("6222 0000 1234 1234");
            payeeEntity.setBankName("中国银行");
            payeeEntity.setName("小明");
            entity.setPayee(payeeEntity);
            list.add(entity);
        }

        Workbook workbook = ExcelExportUtil.exportExcel(params, TemplateExcelExportEntity.class,
            list, map);
        File savefile = new File("d:/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("d:/tt.xls");
        workbook.write(fos);
        fos.close();
    }
}
