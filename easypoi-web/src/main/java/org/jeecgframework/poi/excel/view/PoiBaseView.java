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

import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.jeecgframework.poi.excel.entity.vo.BigExcelConstants;
import org.jeecgframework.poi.excel.entity.vo.MapExcelConstants;
import org.jeecgframework.poi.excel.entity.vo.MapExcelGraphConstants;
import org.jeecgframework.poi.excel.entity.vo.NormalExcelConstants;
import org.jeecgframework.poi.excel.entity.vo.TemplateExcelConstants;
import org.springframework.web.servlet.view.AbstractView;

/**
 *  提供一些通用结构
 */
public abstract class PoiBaseView extends AbstractView {

    protected static boolean isIE(HttpServletRequest request) {
        return (request.getHeader("USER-AGENT").toLowerCase().indexOf("msie") > 0
                || request.getHeader("USER-AGENT").toLowerCase().indexOf("rv:11.0") > 0
                || request.getHeader("USER-AGENT").toLowerCase().indexOf("edge") > 0) ? true
                    : false;
    }

    public static void render(Map<String, Object> model, HttpServletRequest request,
                              HttpServletResponse response, String viewName) throws Exception {
        PoiBaseView view = null;
        if (BigExcelConstants.BIG_EXCEL_VIEW.equals(viewName)) {
            view = new BigExcelExportView();
        } else if (MapExcelConstants.JEECG_MAP_EXCEL_VIEW.equals(viewName)) {
            view = new JeecgMapExcelView();
        } else if (NormalExcelConstants.JEECG_EXCEL_VIEW.equals(viewName)) {
            view = new JeecgSingleExcelView();
        } else if (TemplateExcelConstants.JEECG_TEMPLATE_EXCEL_VIEW.equals(viewName)) {
            view = new JeecgTemplateExcelView();
        } else if (MapExcelGraphConstants.MAP_GRAPH_EXCEL_VIEW.equals(viewName)) {
            view = new MapGraphExcelView();
        }
        view.renderMergedOutputModel(model, request, response);
    }

}
