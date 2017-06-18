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
package cn.afterturn.easypoi.view;

import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.servlet.view.AbstractView;

import cn.afterturn.easypoi.entity.vo.BigExcelConstants;
import cn.afterturn.easypoi.entity.vo.MapExcelConstants;
import cn.afterturn.easypoi.entity.vo.MapExcelGraphConstants;
import cn.afterturn.easypoi.entity.vo.NormalExcelConstants;
import cn.afterturn.easypoi.entity.vo.TemplateExcelConstants;

/**
 *  提供一些通用结构
 */
public abstract class PoiBaseView extends AbstractView {

    private static final Logger LOGGER = LoggerFactory.getLogger(PoiBaseView.class);

    protected static boolean isIE(HttpServletRequest request) {
        return (request.getHeader("USER-AGENT").toLowerCase().indexOf("msie") > 0
                || request.getHeader("USER-AGENT").toLowerCase().indexOf("rv:11.0") > 0
                || request.getHeader("USER-AGENT").toLowerCase().indexOf("edge") > 0) ? true
                    : false;
    }

    public static void render(Map<String, Object> model, HttpServletRequest request,
                              HttpServletResponse response, String viewName) {
        PoiBaseView view = null;
        if (BigExcelConstants.EASYPOI_BIG_EXCEL_VIEW.equals(viewName)) {
            view = new EasypoiBigExcelExportView();
        } else if (MapExcelConstants.EASYPOI_MAP_EXCEL_VIEW.equals(viewName)) {
            view = new EasypoiMapExcelView();
        } else if (NormalExcelConstants.EASYPOI_EXCEL_VIEW.equals(viewName)) {
            view = new EasypoiSingleExcelView();
        } else if (TemplateExcelConstants.EASYPOI_TEMPLATE_EXCEL_VIEW.equals(viewName)) {
            view = new EasypoiTemplateExcelView();
        } else if (MapExcelGraphConstants.MAP_GRAPH_EXCEL_VIEW.equals(viewName)) {
            view = new MapGraphExcelView();
        }
        try {
            view.renderMergedOutputModel(model, request, response);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }
    }

}
