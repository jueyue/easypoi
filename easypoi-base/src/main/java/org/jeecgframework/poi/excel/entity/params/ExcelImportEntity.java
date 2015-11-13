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
package org.jeecgframework.poi.excel.entity.params;

import java.util.List;

/**
 * excel 导入工具类,对cell类型做映射
 * @author JueYue
 * @version 1.0 2013年8月24日
 */
public class ExcelImportEntity extends ExcelBaseEntity {
    /**
     * 对应 Collection NAME
     */
    private String                  collectionName;
    /**
     * 保存图片的地址
     */
    private String                  saveUrl;
    /**
     * 保存图片的类型,1是文件,2是数据库
     */
    private int                     saveType;
    /**
     * 对应exportType
     */
    private String                  classType;
    /**
     * 后缀
     */
    private String                  suffix;
    /**
     * 导入校验字段
     */
    private boolean                 importField;

    private List<ExcelImportEntity> list;

    public String getClassType() {
        return classType;
    }

    public String getCollectionName() {
        return collectionName;
    }

    public List<ExcelImportEntity> getList() {
        return list;
    }

    public int getSaveType() {
        return saveType;
    }

    public String getSaveUrl() {
        return saveUrl;
    }

    public void setClassType(String classType) {
        this.classType = classType;
    }

    public void setCollectionName(String collectionName) {
        this.collectionName = collectionName;
    }

    public void setList(List<ExcelImportEntity> list) {
        this.list = list;
    }

    public void setSaveType(int saveType) {
        this.saveType = saveType;
    }

    public void setSaveUrl(String saveUrl) {
        this.saveUrl = saveUrl;
    }

    public String getSuffix() {
        return suffix;
    }

    public void setSuffix(String suffix) {
        this.suffix = suffix;
    }

    public boolean isImportField() {
        return importField;
    }

    public void setImportField(boolean importField) {
        this.importField = importField;
    }

}
