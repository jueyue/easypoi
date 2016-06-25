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
 * 边框样式
 * @author JueYue
 * 2016年4月3日 上午10:35:28
 */
public class CellStyleBorderEntity {

	private short	borderLeft;
	
	private short	borderRight;
	
	private short	borderTop;
	
	private short	borderBottom;
	
	private short	borderLeftColor;
	
	private short	borderRightColor;
	
	private short	borderTopColor;
	
	private short	borderBottomColor;

	public short getBorderLeft() {
		return borderLeft;
	}

	public void setBorderLeft(short borderLeft) {
		this.borderLeft = borderLeft;
	}

	public short getBorderRight() {
		return borderRight;
	}

	public void setBorderRight(short borderRight) {
		this.borderRight = borderRight;
	}

	public short getBorderTop() {
		return borderTop;
	}

	public void setBorderTop(short borderTop) {
		this.borderTop = borderTop;
	}

	public short getBorderBottom() {
		return borderBottom;
	}

	public void setBorderBottom(short borderBottom) {
		this.borderBottom = borderBottom;
	}

	public short getBorderLeftColor() {
		return borderLeftColor;
	}

	public void setBorderLeftColor(short borderLeftColor) {
		this.borderLeftColor = borderLeftColor;
	}

	public short getBorderRightColor() {
		return borderRightColor;
	}

	public void setBorderRightColor(short borderRightColor) {
		this.borderRightColor = borderRightColor;
	}

	public short getBorderTopColor() {
		return borderTopColor;
	}

	public void setBorderTopColor(short borderTopColor) {
		this.borderTopColor = borderTopColor;
	}

	public short getBorderBottomColor() {
		return borderBottomColor;
	}

	public void setBorderBottomColor(short borderBottomColor) {
		this.borderBottomColor = borderBottomColor;
	}

}
