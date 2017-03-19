package org.jeecgframework.poi.excel.html.css;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.jeecgframework.poi.excel.html.entity.CellStyleBorderEntity;
import org.jeecgframework.poi.excel.html.entity.CellStyleEntity;
import org.jeecgframework.poi.excel.html.entity.HtmlCssConstant;
import org.jeecgframework.poi.util.PoiCssUtils;

/**
 * 把Css样式解析成对应的Model
 * @author JueYue
 * 2017年3月19日
 */
public class CssParseServer {

    public CellStyleEntity parseStyle(String style) {
        Map<String, String> mapStyle = new HashMap<String, String>();
        for (String s : style.split("\\s*;\\s*")) {
            if (StringUtils.isNotBlank(s)) {
                String[] ss = s.split("\\s*\\:\\s*");
                if (ss.length == 2 && StringUtils.isNotBlank(ss[0])
                    && StringUtils.isNotBlank(ss[1])) {
                    String attrName = ss[0].toLowerCase();
                    String attrValue = ss[1];
                    // do not change font name
                    if (!HtmlCssConstant.FONT.equals(attrName)
                        && !HtmlCssConstant.FONT_FAMILY.equals(attrName)) {
                        attrValue = attrValue.toLowerCase();
                    }
                    mapStyle.put(attrName, attrValue);
                }
            }
        }

        return mapToCellStyleEntity(mapStyle);
    }

    private CellStyleEntity mapToCellStyleEntity(Map<String, String> mapStyle) {
        CellStyleEntity entity = new CellStyleEntity();
        entity.setAlign(mapStyle.get(HtmlCssConstant.TEXT_ALIGN));
        entity.setVetical(mapStyle.get(HtmlCssConstant.VETICAL_ALIGN));
        entity.setBackground(getBackground(mapStyle));
        entity.setHeight(mapStyle.get(HtmlCssConstant.HEIGHT));
        entity.setWidth(mapStyle.get(HtmlCssConstant.WIDTH));
        // TODO 这里较为复杂,后期处理
        /*CellStyleBorderEntity border = new CellStyleBorderEntity();
        entity.setBorder(border);
        border.setBorderBottom(borderBottom);*/
        return entity;
    }

    private String getBackground(Map<String, String> style) {
        Map<String, String> mapRtn = new HashMap<String, String>();
        String bg = style.get(HtmlCssConstant.BACKGROUND);
        String bgColor = null;
        if (StringUtils.isNotBlank(bg)) {
            for (String bgAttr : bg.split("(?<=\\)|\\w|%)\\s+(?=\\w)")) {
                if ((bgColor = PoiCssUtils.processColor(bgAttr)) != null) {
                    mapRtn.put(HtmlCssConstant.BACKGROUND_COLOR, bgColor);
                    break;
                }
            }
        }
        bg = style.get(HtmlCssConstant.BACKGROUND_COLOR);
        if (StringUtils.isNotBlank(bg) && (bgColor = PoiCssUtils.processColor(bg)) != null) {
            mapRtn.put(HtmlCssConstant.BACKGROUND_COLOR, bgColor);

        }
        if (bgColor != null) {
            bgColor = mapRtn.get(HtmlCssConstant.BACKGROUND_COLOR);
            if (!"#ffffff".equals(bgColor)) {
                return mapRtn.get(HtmlCssConstant.BACKGROUND_COLOR);
            }
        }
        return null;
    }

}
