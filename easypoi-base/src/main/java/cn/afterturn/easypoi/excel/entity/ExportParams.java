/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package cn.afterturn.easypoi.excel.entity;

import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.export.styler.ExcelExportStylerDefaultImpl;
import lombok.Data;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * Excel 导出参数
 *
 * @author JueYue
 * @version 1.0 2013年8月24日
 */
@Data
public class ExportParams extends ExcelBaseParams {

    /**
     * 表格名称
     */
    private String title;

    /**
     * 表格名称
     */
    private short titleHeight = 10;

    /**
     * 第二行名称
     */
    private String secondTitle;

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
     * 表头颜色 &  标题颜色
     */
    private short     color             = HSSFColor.HSSFColorPredefined.WHITE.getIndex();
    /**
     * 第二行标题颜色
     * 属性说明行的颜色 例如:HSSFColor.SKY_BLUE.index 默认
     */
    private short     headerColor       = HSSFColor.HSSFColorPredefined.SKY_BLUE.getIndex();
    /**
     * Excel 导出版本
     */
    private ExcelType type              = ExcelType.HSSF;
    /**
     * Excel 导出style
     */
    private Class<?>  style             = ExcelExportStylerDefaultImpl.class;

    /**
     * 表头高度
     */
    private double  headerHeight     = 9D;
    /**
     * 是否创建表头
     */
    private boolean isCreateHeadRows = true;
    /**
     * 是否动态获取数据
     */
    private boolean isDynamicData    = false;
    /**
     * 是否追加图形
     */
    private boolean isAppendGraph    = true;
    /**
     * 是否固定表头
     */
    private boolean isFixedTitle     = true;
    /**
     * 单sheet最大值
     * 03版本默认6W行,07默认100W
     */
    private int     maxNum           = 0;

    /**
     * 导出时在excel中每个列的高度 单位为字符，一个汉字=2个字符
     * 全局设置,优先使用
     */
    private short height = 0;

    /**
     * 只读
     */
    private boolean readonly = false;

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

}
