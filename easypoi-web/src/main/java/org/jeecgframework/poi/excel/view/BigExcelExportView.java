/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 *   
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.jeecgframework.poi.excel.view;

import java.util.Collections;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.vo.BigExcelConstants;
import org.jeecgframework.poi.handler.inter.IExcelExportServer;
import org.springframework.stereotype.Controller;

/**
 * @author JueYue on 14-3-8. Excel 生成解析器,减少用户操作
 */
@Controller(BigExcelConstants.BIG_EXCEL_VIEW)
public class BigExcelExportView extends MiniAbstractExcelView {

    public BigExcelExportView() {
        super();
    }

    @Override
    protected void renderMergedOutputModel(Map<String, Object> model, HttpServletRequest request,
                                           HttpServletResponse response) throws Exception {
        String codedFileName = "临时文件";
        Workbook workbook = ExcelExportUtil.exportBigExcel(
            (ExportParams) model.get(BigExcelConstants.PARAMS),
            (Class<?>) model.get(BigExcelConstants.CLASS), Collections.EMPTY_LIST);
        IExcelExportServer server = (IExcelExportServer) model.get(BigExcelConstants.DATA_INTER);
        int page = 1;
        List<Object> list = server
            .selectListForExcelExport(model.get(BigExcelConstants.DATA_PARAMS), page++);
        while (list != null && list.size() > 0) {
            workbook = ExcelExportUtil.exportBigExcel(
                (ExportParams) model.get(BigExcelConstants.PARAMS),
                (Class<?>) model.get(BigExcelConstants.CLASS), list);
            list = server.selectListForExcelExport(model.get(BigExcelConstants.DATA_PARAMS),
                page++);
        }
        ExcelExportUtil.closeExportBigExcel();
        if (model.containsKey(BigExcelConstants.FILE_NAME)) {
            codedFileName = (String) model.get(BigExcelConstants.FILE_NAME);
        }
        if (workbook instanceof HSSFWorkbook) {
            codedFileName += HSSF;
        } else {
            codedFileName += XSSF;
        }
        if (isIE(request)) {
            codedFileName = java.net.URLEncoder.encode(codedFileName, "UTF8");
        } else {
            codedFileName = new String(codedFileName.getBytes("UTF-8"), "ISO-8859-1");
        }
        response.setHeader("content-disposition", "attachment;filename=" + codedFileName);
        ServletOutputStream out = response.getOutputStream();
        workbook.write(out);
        out.flush();
    }
}
