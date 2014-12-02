package org.jeecgframework.poi.excel.view;

import java.util.Collection;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.entity.vo.MapExcelConstants;
import org.jeecgframework.poi.excel.export.ExcelExportServer;
import org.springframework.web.servlet.view.document.AbstractExcelView;

/**
 * Map 对象接口
 * 
 * @author JueYue
 * @date 2014年11月25日 下午3:26:32
 */
@SuppressWarnings("unchecked")
public class JeecgMapExcelView extends AbstractExcelView {

    @Override
    protected void buildExcelDocument(Map<String, Object> model, HSSFWorkbook hssfWorkbook,
                                      HttpServletRequest httpServletRequest,
                                      HttpServletResponse httpServletResponse) throws Exception {
        String codedFileName = "临时文件.xls";
        if (model.containsKey(MapExcelConstants.FILE_NAME)) {
            codedFileName = (String) model.get(MapExcelConstants.FILE_NAME) + ".xls";
        }
        if (isIE(httpServletRequest)) {
            codedFileName = java.net.URLEncoder.encode(codedFileName, "UTF8");
        } else {
            codedFileName = new String(codedFileName.getBytes("UTF-8"), "ISO-8859-1");
        }
        httpServletResponse
            .setHeader("content-disposition", "attachment;filename=" + codedFileName);
        new ExcelExportServer().createSheetForMap(hssfWorkbook,
            (ExportParams) model.get(MapExcelConstants.PARAMS),
            (List<ExcelExportEntity>) model.get(MapExcelConstants.ENTITY_LIST),
            (Collection<? extends Map<?, ?>>) model.get(MapExcelConstants.MAP_LIST),
            ((ExportParams) model.get(MapExcelConstants.PARAMS)).getType());
    }

    public boolean isIE(HttpServletRequest request) {
        return (request.getHeader("USER-AGENT").toLowerCase().indexOf("msie") > 0 || request
            .getHeader("USER-AGENT").toLowerCase().indexOf("rv:11.0") > 0) ? true : false;
    }
}
