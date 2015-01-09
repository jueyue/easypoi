package org.jeecgframework.poi.excel.entity.enmus;

import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerBorderImpl;
import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerColorImpl;
import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerDefaultImpl;

/**
 * 插件提供的几个默认样式
 * @author JueYue
 * @date 2015年1月9日 下午9:02:24
 */
public enum ExcelStyleType {

    NONE("默认样式", ExcelExportStylerDefaultImpl.class),
    BORDER("边框样式", ExcelExportStylerBorderImpl.class),
    COLOR("间隔行样式", ExcelExportStylerColorImpl.class);

    private String   name;
    private Class<?> clazz;

    ExcelStyleType(String name, Class<?> clazz) {
        this.name = name;
        this.clazz = clazz;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public String getName() {
        return name;
    }

}
