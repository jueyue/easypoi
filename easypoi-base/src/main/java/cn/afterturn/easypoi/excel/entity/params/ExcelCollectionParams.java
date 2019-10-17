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
package cn.afterturn.easypoi.excel.entity.params;

import lombok.Data;

import java.util.Map;

/**
 * Excel 对于的 Collection
 * 
 * @author JueYue
 *  2013-9-26
 * @version 1.0
 */
@Data
public class ExcelCollectionParams {

    /**
     * 集合对应的名称
     */
    private String                         name;
    /**
     * Excel 列名称
     */
    private String                         excelName;
    /**
     * 实体对象
     */
    private Class<?>                       type;
    /**
     * 这个list下面的参数集合实体对象
     */
    private Map<String, ExcelImportEntity> excelParams;

}
