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
package org.jeecgframework.poi.excel.entity;

import org.apache.poi.hssf.util.HSSFColor;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerDefaultImpl;

/**
 * Excel 导出参数
 * 
 * @author JueYue
 * @version 1.0 2013年8月24日
 */
public class ExportParams extends ExcelBaseParams {

    /**
     * 表格名称
     */
    private String    title;

    /**
     * 表格名称
     */
    private short     titleHeight       = 10;

    /**
     * 第二行名称
     */
    private String    secondTitle;

    /**
     * 表格名称
     */
    private short     secondTitleHeight = 8;
    /**
     * sheetName
     */
    private String    sheetName;
    /**
     * 过滤的属性
     */
    private String[]  exclusions;
    /**
     * 是否添加需要需要
     */
    private boolean   addIndex;
    /**
     * 是否添加需要需要
     */
    private String    indexName         = "序号";
    /**
     * 冰冻列
     */
    private int       freezeCol;
    /**
     * 表头颜色
     */
    private short     color             = HSSFColor.WHITE.index;
    /**
     * 属性说明行的颜色 例如:HSSFColor.SKY_BLUE.index 默认
     */
    private short     headerColor       = HSSFColor.SKY_BLUE.index;
    /**
     * Excel 导出版本
     */
    private ExcelType type              = ExcelType.HSSF;
    /**
     * Excel 导出style
     */
    private Class<?>  style             = ExcelExportStylerDefaultImpl.class;
    /**
     * 是否创建表头
     */
    private boolean   isCreateHeadRows  = true;
    /**
     * 是否动态获取数据
     */
    private boolean   isDynamicData     = false;
    /**
     * 是否追加图形
     */
    private boolean   isAppendGraph     = true;
    /**
     * 单sheet最大值
     * 03版本默认6W行,07默认100W
     */
    private int       maxNum            = 0;

    public ExportParams() {

    }

    public ExportParams(String title, String sheetName) {
        this.title = title;
        this.sheetName = sheetName;
    }

    public ExportParams(String title, String sheetName, ExcelType type) {
        this.title = title;
        this.sheetName = sheetName;
        this.type = type;
    }

    public ExportParams(String title, String secondTitle, String sheetName) {
        this.title = title;
        this.secondTitle = secondTitle;
        this.sheetName = sheetName;
    }

    public short getColor() {
        return color;
    }

    public String[] getExclusions() {
        return exclusions;
    }

    public short getHeaderColor() {
        return headerColor;
    }

    public String getSecondTitle() {
        return secondTitle;
    }

    public short getSecondTitleHeight() {
        return (short) (secondTitleHeight * 50);
    }

    public String getSheetName() {
        return sheetName;
    }

    public String getTitle() {
        return title;
    }

    public short getTitleHeight() {
        return (short) (titleHeight * 50);
    }

    public boolean isAddIndex() {
        return addIndex;
    }

    public void setAddIndex(boolean addIndex) {
        this.addIndex = addIndex;
    }

    public void setColor(short color) {
        this.color = color;
    }

    public void setExclusions(String[] exclusions) {
        this.exclusions = exclusions;
    }

    public void setHeaderColor(short headerColor) {
        this.headerColor = headerColor;
    }

    public void setSecondTitle(String secondTitle) {
        this.secondTitle = secondTitle;
    }

    public void setSecondTitleHeight(short secondTitleHeight) {
        this.secondTitleHeight = secondTitleHeight;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public void setTitleHeight(short titleHeight) {
        this.titleHeight = titleHeight;
    }

    public ExcelType getType() {
        return type;
    }

    public void setType(ExcelType type) {
        this.type = type;
    }

    public String getIndexName() {
        return indexName;
    }

    public void setIndexName(String indexName) {
        this.indexName = indexName;
    }

    public Class<?> getStyle() {
        return style;
    }

    public void setStyle(Class<?> style) {
        this.style = style;
    }

    public int getFreezeCol() {
        return freezeCol;
    }

    public void setFreezeCol(int freezeCol) {
        this.freezeCol = freezeCol;
    }

    public boolean isCreateHeadRows() {
        return isCreateHeadRows;
    }

    public void setCreateHeadRows(boolean isCreateHeadRows) {
        this.isCreateHeadRows = isCreateHeadRows;
    }

    public boolean isDynamicData() {
        return isDynamicData;
    }

    public void setDynamicData(boolean isDynamicData) {
        this.isDynamicData = isDynamicData;
    }

    public boolean isAppendGraph() {
        return isAppendGraph;
    }

    public void setAppendGraph(boolean isAppendGraph) {
        this.isAppendGraph = isAppendGraph;
    }

	public int getMaxNum() {
		return maxNum;
	}

	public void setMaxNum(int maxNum) {
		this.maxNum = maxNum;
	}

}
