package org.jeecgframework.poi.excel.view;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.vo.NormalExcelConstants;
import org.jeecgframework.poi.excel.export.ExcelExportServer;
import org.springframework.web.servlet.view.document.AbstractExcelView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * @Author JueYue on 14-3-8. Excel 生成解析器,减少用户操作
 */
@SuppressWarnings("unchecked")
public class JeecgSingleExcelView extends AbstractExcelView {

	@Override
	protected void buildExcelDocument(Map<String, Object> model,
			HSSFWorkbook hssfWorkbook, HttpServletRequest httpServletRequest,
			HttpServletResponse httpServletResponse) throws Exception {
		String codedFileName = "临时文件.xls";
		if (model.containsKey(NormalExcelConstants.FILE_NAME)) {
			codedFileName = (String) model.get(NormalExcelConstants.FILE_NAME)
					+ ".xls";
		}
		httpServletResponse.setHeader("content-disposition",
				"attachment;filename="
						+ new String(codedFileName.getBytes(), "iso8859-1"));
		if (model.containsKey(NormalExcelConstants.MAP_LIST)) {
			List<Map<String, Object>> list = (List<Map<String, Object>>) model
					.get(NormalExcelConstants.MAP_LIST);
			for (Map<String, Object> map : list) {
				new ExcelExportServer().createSheet(hssfWorkbook,
						(ExportParams) map.get(NormalExcelConstants.PARAMS),
						(Class<?>) map.get(NormalExcelConstants.CLASS),
						(Collection<?>) map.get(NormalExcelConstants.DATA_LIST));
			}
		} else {
			new ExcelExportServer().createSheet(hssfWorkbook,
					(ExportParams) model.get(NormalExcelConstants.PARAMS),
					(Class<?>) model.get(NormalExcelConstants.CLASS),
					(Collection<?>) model.get(NormalExcelConstants.DATA_LIST));
		}
	}
}
