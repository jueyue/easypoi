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
package cn.afterturn.easypoi.entity;

/**
 * word导出,图片设置和图片信息
 * 
 * @author JueYue
 *  2017-9-15
 * @version 3.0.1
 */
public class ImageEntity {

    public static String URL  = "url";
    public static String Data = "data";
    /**
     * 图片输入方式
     */
    private String       type = URL;
    /**
     * 图片宽度
     */
    private int          width;
    // 图片高度
    private int          height;
    // 图片地址
    private String       url;
    // 图片信息
    private byte[]       data;

    private int          rowspan = 1;
    private int          colspan = 1;


    public ImageEntity() {

    }

    public ImageEntity(byte[] data, int width, int height) {
        this.data = data;
        this.width = width;
        this.height = height;
        this.type = Data;
    }

    public ImageEntity(String url, int width, int height) {
        this.url = url;
        this.width = width;
        this.height = height;
    }

    public byte[] getData() {
        return data;
    }

    public int getHeight() {
        return height;
    }

    public String getType() {
        return type;
    }

    public String getUrl() {
        return url;
    }

    public int getWidth() {
        return width;
    }

    public void setData(byte[] data) {
        this.data = data;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public void setType(String type) {
        this.type = type;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public int getRowspan() {
        return rowspan;
    }

    public void setRowspan(int rowspan) {
        this.rowspan = rowspan;
    }

    public int getColspan() {
        return colspan;
    }

    public void setColspan(int colspan) {
        this.colspan = colspan;
    }
}
