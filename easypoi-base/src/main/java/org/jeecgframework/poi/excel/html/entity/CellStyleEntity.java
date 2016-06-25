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
package org.jeecgframework.poi.excel.html.entity;

/**
 * Cell 具有的样式
 * @author JueYue
 * 2016年3月20日 下午4:56:51
 */
public class CellStyleEntity {
	/**
	 * 宽
	 */
	private double					width;
	/**
	 * 高
	 */
	private double					height;
	/**
	 * 边框
	 */
	private CellStyleBorderEntity	border;
	/**
	 * 背景
	 */
	private String					background;
	/**
	 * 水平位置
	 */
	private String					align;
	/**
	 * 垂直位置
	 */
	private String					vetical;

	public double getWidth() {
		return width;
	}

	public void setWidth(double width) {
		this.width = width;
	}

	public double getHeight() {
		return height;
	}

	public void setHeight(double height) {
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

}
