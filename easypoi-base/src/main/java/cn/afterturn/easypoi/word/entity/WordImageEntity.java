/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License. You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under the License
 * is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
 * or implied. See the License for the specific language governing permissions and limitations under
 * the License.
 */
package cn.afterturn.easypoi.word.entity;

import cn.afterturn.easypoi.entity.ImageEntity;

/**
 * word导出,图片设置和图片信息
 *
 * @author JueYue
 *  2013-11-17
 * @version 1.0
 */
@Deprecated
public class WordImageEntity extends ImageEntity {

    public WordImageEntity() {

    }

    public WordImageEntity(byte[] data, int width, int height) {
        super(data, width, height);
    }

    public WordImageEntity(String url, int width, int height) {
        super(url, width, height);
    }

}
