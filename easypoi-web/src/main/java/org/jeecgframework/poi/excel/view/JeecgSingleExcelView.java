package org.jeecgframework.poi.excel.view;

import java.util.Collection;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.vo.MapExcelConstants;
import org.jeecgframework.poi.excel.entity.vo.NormalExcelConstants;
import org.jeecgframework.poi.excel.export.ExcelExportServer;
import org.springframework.web.servlet.view.document.AbstractExcelView;

/**
 * @Author JueYue on 14-3-8. Excel 生成解析器,减少用户操作
 */
@SuppressWarnings("unchecked")
public class JeecgSingleExcelView extends AbstractExcelView {

    @Override
    protected void buildExcelDocument(Map<String, Object> model, HSSFWorkbook hssfWorkbook,
                                      HttpServletRequest httpServletRequest,
                                      HttpServletResponse httpServletResponse) throws Exception {
        String codedFileName = "临时文件.xls";
        if (model.containsKey(NormalExcelConstants.FILE_NAME)) {
            codedFileName = (String) model.get(NormalExcelConstants.FILE_NAME) + ".xls";
        }
        if (isIE(httpServletRequest)) {
            codedFileName = java.net.URLEncoder.encode(codedFileName, "UTF8");
        } else {
            codedFileName = new String(codedFileName.getBytes("UTF-8"), "ISO-8859-1");
        }
        httpServletResponse
            .setHeader("content-disposition", "attachment;filename=" + codedFileName);
        if (model.containsKey(NormalExcelConstants.MAP_LIST)) {
            List<Map<String, Object>> list = (List<Map<String, Object>>) model
                .get(NormalExcelConstants.MAP_LIST);
            for (Map<String, Object> map : list) {
                new ExcelExportServer().createSheet(hssfWorkbook,
                    (ExportParams) map.get(NormalExcelConstants.PARAMS),
                    (Class<?>) map.get(NormalExcelConstants.CLASS),
                    (Collection<?>) map.get(NormalExcelConstants.DATA_LIST),
                    ((ExportParams) model.get(MapExcelConstants.PARAMS)).getType());
            }
        } else {
            new ExcelExportServer().createSheet(hssfWorkbook,
                (ExportParams) model.get(NormalExcelConstants.PARAMS),
                (Class<?>) model.get(NormalExcelConstants.CLASS),
                (Collection<?>) model.get(NormalExcelConstants.DATA_LIST),
                ((ExportParams) model.get(MapExcelConstants.PARAMS)).getType());
        }
    }

    public boolean isIE(HttpServletRequest request) {
        return (request.getHeader("USER-AGENT").toLowerCase().indexOf("msie") > 0 || request
            .getHeader("USER-AGENT").toLowerCase().indexOf("rv:11.0") > 0) ? true : false;
    }
}
