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
package cn.aftertrun.easypoi.excel.html.entity.style;

/**
 * Cell 具有的样式
 * @author JueYue
 * 2016年3月20日 下午4:56:51
 */
public class CellStyleEntity {
    /**
     * 宽
     */
    private String                width;
    /**
     * 高
     */
    private String                height;
    /**
     * 边框
     */
    private CellStyleBorderEntity border;
    /**
     * 背景
     */
    private String                background;
    /**
     * 水平位置
     */
    private String                align;
    /**
     * 垂直位置
     */
    private String                vetical;
    /**
     * 字体设置
     */
    private CssStyleFontEnity     font;

    public String getWidth() {
        return width;
    }

    public void setWidth(String width) {
        this.width = width;
    }

    public String getHeight() {
        return height;
    }

    public void setHeight(String height) {
        this.height = height;
    }

    public CellStyleBorderEntity getBorder() {
        return border;
    }

    public void setBorder(CellStyleBorderEntity border) {
        this.border = border;
    }

    public String getBackground() {
        return background;
    }

    public void setBackground(String background) {
        this.background = background;
    }

    public String getAlign() {
        return align;
    }

    public void setAlign(String align) {
        this.align = align;
    }

    public String getVetical() {
        return vetical;
    }

    public void setVetical(String vetical) {
        this.vetical = vetical;
    }

    public CssStyleFontEnity getFont() {
        return font;
    }

    public void setFont(CssStyleFontEnity font) {
        this.font = font;
    }

    @Override
    public String toString() {
        return new StringBuilder().append(align).append(background).append(border).append(height)
            .append(vetical).append(width).append(font).toString();
    }

}
