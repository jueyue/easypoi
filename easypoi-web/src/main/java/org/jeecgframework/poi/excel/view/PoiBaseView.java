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

import javax.servlet.http.HttpServletRequest;

import org.springframework.web.servlet.view.AbstractView;

/**
 *  提供一些通用结构
 */
public abstract class PoiBaseView extends AbstractView {
    
    
    protected boolean isIE(HttpServletRequest request) {
        return (request.getHeader("USER-AGENT").toLowerCase().indexOf("msie") > 0
                || request.getHeader("USER-AGENT").toLowerCase().indexOf("rv:11.0") > 0
                || request.getHeader("USER-AGENT").toLowerCase().indexOf("edge") > 0) ? true
                    : false;
    }

}
