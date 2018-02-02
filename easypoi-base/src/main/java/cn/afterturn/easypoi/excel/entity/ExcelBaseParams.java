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
package cn.afterturn.easypoi.excel.entity;

import cn.afterturn.easypoi.handler.inter.IExcelDataHandler;
import cn.afterturn.easypoi.handler.inter.IExcelDictHandler;

/**
 * 基础参数
 * @author JueYue
 *  2014年6月20日 下午1:56:52
 */
@SuppressWarnings("rawtypes")
public class ExcelBaseParams {

    /**
     * 数据处理接口,以此为主,replace,format都在这后面
     */
    private IExcelDataHandler dataHandler;

    /**
     * 字段处理类
     */
    private IExcelDictHandler dictHandler;

    public IExcelDataHandler getDataHandler() {
        return dataHandler;
    }

    public void setDataHandler(IExcelDataHandler dataHandler) {
        this.dataHandler = dataHandler;
    }

    public IExcelDictHandler getDictHandler() {
        return dictHandler;
    }

    public void setDictHandler(IExcelDictHandler dictHandler) {
        this.dictHandler = dictHandler;
    }
}
