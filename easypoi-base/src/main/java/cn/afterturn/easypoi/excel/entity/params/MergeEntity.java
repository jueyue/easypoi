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

import java.util.List;

/**
 * 合并单元格使用对象
 *
 * Created by jue on 14-6-11.
 */
@Data
public class MergeEntity {
    /**
     * 合并开始行
     */
    private int          startRow;
    /**
     * 合并结束行
     */
    private int          endRow;
    /**
     * 文字
     */
    private String       text;
    /**
     * 依赖关系文本
     */
    private List<String> relyList;

    public MergeEntity() {

    }

    public MergeEntity(String text, int startRow, int endRow) {
        this.text = text;
        this.endRow = endRow;
        this.startRow = startRow;
    }

}
