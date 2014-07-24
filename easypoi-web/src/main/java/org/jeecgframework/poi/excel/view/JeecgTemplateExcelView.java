package org.jeecgframework.poi.excel.view;

import java.util.List;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.jeecgframework.poi.excel.entity.vo.NormalPOIConstants;
import org.jeecgframework.poi.excel.entity.vo.TemplatePOIConstants;
import org.springframework.web.servlet.view.AbstractView;

/**
 * 模板视图
 * 
 * @author JueYue
 * @date 2014年6月30日 下午9:15:49
 */
@SuppressWarnings("unchecked")
public class JeecgTemplateExcelView extends AbstractView {

	private static final String CONTENT_TYPE = "application/vnd.ms-excel";

	public JeecgTemplateExcelView() {
		setContentType(CONTENT_TYPE);
	}

	@Override
	protected void renderMergedOutputModel(Map<String, Object> model,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String codedFileName = "临时文件.xls";
		if (model.containsKey(NormalPOIConstants.FILE_NAME)) {
			codedFileName = (String) model.get(NormalPOIConstants.FILE_NAME)
					+ ".xls";
		}
		response.setHeader("content-disposition", "attachment;filename="
				+ new String(codedFileName.getBytes(), "iso8859-1"));
		Workbook workbook = ExcelExportUtil.exportExcel(
				(TemplateExportParams)model.get(TemplatePOIConstants.PARAMS),
				(Class<?>) model.get(TemplatePOIConstants.CLASS),
				(List<?>)model.get(TemplatePOIConstants.LIST_DATA),
				(Map<String, Object>)model.get(TemplatePOIConstants.MAP_DATA));
		ServletOutputStream out = response.getOutputStream();
		workbook.write(out);
		out.flush();
	}
}
