package org.jeecgframework.poi.excel.view;

import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jeecgframework.poi.excel.entity.vo.TemplateWordConstants;
import org.jeecgframework.poi.word.WordExportUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.servlet.view.AbstractView;

/**
 * Word模板视图
 * 
 * @author JueYue
 * @date 2014年6月30日 下午9:15:49
 */
@SuppressWarnings("unchecked")
@Controller(TemplateWordConstants.JEECG_TEMPLATE_WORD_VIEW)
public class JeecgTemplateWordView extends AbstractView {

    private static final String CONTENT_TYPE = "application/msword";

    public JeecgTemplateWordView() {
        setContentType(CONTENT_TYPE);
    }

    public boolean isIE(HttpServletRequest request) {
        return (request.getHeader("USER-AGENT").toLowerCase().indexOf("msie") > 0 || request
            .getHeader("USER-AGENT").toLowerCase().indexOf("rv:11.0") > 0) ? true : false;
    }

    @Override
    protected void renderMergedOutputModel(Map<String, Object> model, HttpServletRequest request,
                                           HttpServletResponse response) throws Exception {
        String codedFileName = "临时文件.docx";
        if (model.containsKey(TemplateWordConstants.FILE_NAME)) {
            codedFileName = (String) model.get(TemplateWordConstants.FILE_NAME) + ".docx";
        }
        if (isIE(request)) {
            codedFileName = java.net.URLEncoder.encode(codedFileName, "UTF8");
        } else {
            codedFileName = new String(codedFileName.getBytes("UTF-8"), "ISO-8859-1");
        }
        response.setHeader("content-disposition", "attachment;filename=" + codedFileName);
        XWPFDocument document = WordExportUtil.exportWord07(
            (String) model.get(TemplateWordConstants.URL),
            (Map<String, Object>) model.get(TemplateWordConstants.MAP_DATA));
        ServletOutputStream out = response.getOutputStream();
        document.write(out);
        out.flush();
    }
}
