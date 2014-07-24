package org.jeecgframework.poi.excel.view;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.vo.NormalPOIConstants;
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
		if (model.containsKey(NormalPOIConstants.FILE_NAME)) {
			codedFileName = (String) model.get(NormalPOIConstants.FILE_NAME)
					+ ".xls";
		}
		httpServletResponse.setHeader("content-disposition",
				"attachment;filename="
						+ new String(codedFileName.getBytes(), "iso8859-1"));
		if (model.containsKey(NormalPOIConstants.MAP_LIST)) {
			List<Map<String, Object>> list = (List<Map<String, Object>>) model
					.get(NormalPOIConstants.MAP_LIST);
			for (Map<String, Object> map : list) {
				new ExcelExportServer().createSheet(hssfWorkbook,
						(ExportParams) map.get(NormalPOIConstants.PARAMS),
						(Class<?>) map.get(NormalPOIConstants.CLASS),
						(Collection<?>) map.get(NormalPOIConstants.DATA_LIST));
			}
		} else {
			new ExcelExportServer().createSheet(hssfWorkbook,
					(ExportParams) model.get(NormalPOIConstants.PARAMS),
					(Class<?>) model.get(NormalPOIConstants.CLASS),
					(Collection<?>) model.get(NormalPOIConstants.DATA_LIST));
		}
	}
}
