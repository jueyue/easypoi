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
package cn.afterturn.easypoi.word.entity.params;

import java.util.List;

import cn.afterturn.easypoi.excel.entity.ExcelBaseParams;
import cn.afterturn.easypoi.handler.inter.IExcelDataHandler;

/**
 * Excel 导出对象
 * 
 * @author JueYue
 *  2014年8月9日 下午10:21:13
 */
public class ExcelListEntity extends ExcelBaseParams {

    /**
     * 数据源
     */
    private List<?>  list;

    /**
     * 实体类对象
     */
    private Class<?> clazz;

    /**
     * 表头行数
     */
    private int      headRows = 1;

    public ExcelListEntity() {

    }

    public ExcelListEntity(List<?> list, Class<?> clazz) {
        this.list = list;
        this.clazz = clazz;
    }

    public ExcelListEntity(List<?> list, Class<?> clazz, IExcelDataHandler dataHandler) {
        this.list = list;
        this.clazz = clazz;
        setDataHandler(dataHandler);
    }

    public ExcelListEntity(List<?> list, Class<?> clazz, IExcelDataHandler dataHandler,
                           int headRows) {
        this.list = list;
        this.clazz = clazz;
        this.headRows = headRows;
        setDataHandler(dataHandler);
    }

    public ExcelListEntity(List<?> list, Class<?> clazz, int headRows) {
        this.list = list;
        this.clazz = clazz;
        this.headRows = headRows;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public int getHeadRows() {
        return headRows;
    }

    public List<?> getList() {
        return list;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
    }

    public void setHeadRows(int headRows) {
        this.headRows = headRows;
    }

    public void setList(List<?> list) {
        this.list = list;
    }

}
