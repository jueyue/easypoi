package org.jeecgframework.poi.excel.html.css;

import static org.jeecgframework.poi.excel.html.entity.HtmlCssConstant.*;

import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.jeecgframework.poi.excel.html.entity.style.CellStyleEntity;
import org.jeecgframework.poi.excel.html.entity.style.CssStyleFontEnity;
import org.jeecgframework.poi.util.PoiCssUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 把Css样式解析成对应的Model
 * @author JueYue
 * 2017年3月19日
 */
public class CssParseServer {

    private static final Logger log = LoggerFactory.getLogger(CssParseServer.class);

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
                    if (!FONT.equals(attrName) && !FONT_FAMILY.equals(attrName)) {
                        attrValue = attrValue.toLowerCase();
                    }
                    mapStyle.put(attrName, attrValue);
                }
            }
        }

        parseFontAttr(mapStyle);

        return mapToCellStyleEntity(mapStyle);
    }

    private void parseFontAttr(Map<String, String> mapRtn) {

        log.debug("Parse Font Style.");
        // color
        String color = PoiCssUtils.processColor(mapRtn.get(COLOR));
        if (StringUtils.isNotBlank(color)) {
            log.debug("Text Color [{}] Found.", color);
            mapRtn.put(COLOR, color);
        }
        // font
        String font = mapRtn.get(FONT);
        if (StringUtils.isNotBlank(font)
            && !ArrayUtils.contains(new String[] { "small-caps", "caption", "icon", "menu",
                                                   "message-box", "small-caption", "status-bar" },
                font)) {
            log.debug("Parse Font Attr [{}].", font);
            String[] ignoreStyles = new String[] { "normal",
                                                   // font weight normal
                                                   "[1-3]00" };
            StringBuffer sbFont = new StringBuffer(
                font.replaceAll("^|\\s*" + StringUtils.join(ignoreStyles, "|") + "\\s+|$", " "));
            log.debug("Font Attr [{}] After Process Ingore.", sbFont);
            // style
            Matcher m = Pattern.compile("(?:^|\\s+)(italic|oblique)(?:\\s+|$)")
                .matcher(sbFont.toString());
            if (m.find()) {
                sbFont.setLength(0);
                if (log.isDebugEnabled()) {
                    log.debug("Font Style [{}] Found.", m.group(1));
                }
                mapRtn.put(FONT_STYLE, ITALIC);
                m.appendReplacement(sbFont, " ");
                m.appendTail(sbFont);
            }
            // weight
            m = Pattern.compile("(?:^|\\s+)(bold(?:er)?|[7-9]00)(?:\\s+|$)")
                .matcher(sbFont.toString());
            if (m.find()) {
                sbFont.setLength(0);
                if (log.isDebugEnabled()) {
                    log.debug("Font Weight [{}](bold) Found.", m.group(1));
                }
                mapRtn.put(FONT_WEIGHT, BOLD);
                m.appendReplacement(sbFont, " ");
                m.appendTail(sbFont);
            }
            // size xx-small | x-small | small | medium | large | x-large | xx-large | 18px [/2]
            m = Pattern.compile(
                // before blank or start
                new StringBuilder("(?:^|\\s+)")
                    // font size
                    .append("(xx-small|x-small|small|medium|large|x-large|xx-large|").append("(?:")
                    .append(PATTERN_LENGTH).append("))")
                    // line height
                    .append("(?:\\s*\\/\\s*(").append(PATTERN_LENGTH).append("))?")
                    // after blank or end
                    .append("(?:\\s+|$)").toString())
                .matcher(sbFont.toString());
            if (m.find()) {
                sbFont.setLength(0);
                log.debug("Font Size[/line-height] [{}] Found.", m.group());
                String fontSize = m.group(1);
                if (StringUtils.isNotBlank(fontSize)) {
                    fontSize = StringUtils.deleteWhitespace(fontSize);
                    log.debug("Font Size [{}].", fontSize);
                    if (fontSize.matches(PATTERN_LENGTH)) {
                        mapRtn.put(FONT_SIZE, fontSize);
                    } else {
                        log.info("Font Size [{}] Not Supported, Ignore.", fontSize);
                    }
                }
                String lineHeight = m.group(2);
                if (StringUtils.isNotBlank(lineHeight)) {
                    log.info("Line Height [{}] Not Supported, Ignore.", lineHeight);
                }
                m.appendReplacement(sbFont, " ");
                m.appendTail(sbFont);
            }
            // font family
            if (sbFont.length() > 0) {
                log.debug("Font Families [{}].", sbFont);
                // trim & remove '"
                String fontFamily = sbFont.toString().split("\\s*,\\s*")[0].trim()
                    .replaceAll("'|\"", "");
                log.debug("Use First Font Family [{}].", fontFamily);
                mapRtn.put(FONT_FAMILY, fontFamily);
            }
        }
        font = mapRtn.get(FONT_STYLE);
        if (ArrayUtils.contains(new String[] { ITALIC, "oblique" }, font)) {
            log.debug("Font Italic [{}] Found.", font);
            mapRtn.put(FONT_STYLE, ITALIC);
        }
        font = mapRtn.get(FONT_WEIGHT);
        if (StringUtils.isNotBlank(font) && Pattern.matches("^bold(?:er)?|[7-9]00$", font)) {
            log.debug("Font Weight [{}](bold) Found.", font);
            mapRtn.put(FONT_WEIGHT, BOLD);
        }
        font = mapRtn.get(FONT_SIZE);
        if (!PoiCssUtils.isNum(font)) {
            log.debug("Font Size [{}] Error.", font);
            mapRtn.remove(FONT_SIZE);
        }
    }

    private CellStyleEntity mapToCellStyleEntity(Map<String, String> mapStyle) {
        CellStyleEntity entity = new CellStyleEntity();
        entity.setAlign(mapStyle.get(TEXT_ALIGN));
        entity.setVetical(mapStyle.get(VETICAL_ALIGN));
        entity.setBackground(getBackground(mapStyle));
        entity.setHeight(mapStyle.get(HEIGHT));
        entity.setWidth(mapStyle.get(WIDTH));
        entity.setFont(getCssStyleFontEnity(mapStyle));
        // TODO 这里较为复杂,后期处理
        /*CellStyleBorderEntity border = new CellStyleBorderEntity();
        entity.setBorder(border);
        border.setBorderBottom(borderBottom);*/
        return entity;
    }

    private CssStyleFontEnity getCssStyleFontEnity(Map<String, String> style) {
        CssStyleFontEnity font = new CssStyleFontEnity();
        font.setStyle(style.get(FONT_STYLE));
        int fontSize = PoiCssUtils.getInt(style.get(FONT_SIZE));
        if (fontSize > 0) {
            font.setSize(fontSize);
        }
        font.setWeight(style.get(FONT_WEIGHT));
        font.setFamily(style.get(FONT_FAMILY));
        font.setDecoration(style.get(TEXT_DECORATION));
        font.setColor(PoiCssUtils.parseColor(style.get(COLOR)));
        return font;
    }

    private String getBackground(Map<String, String> style) {
        Map<String, String> mapRtn = new HashMap<String, String>();
        String bg = style.get(BACKGROUND);
        String bgColor = null;
        if (StringUtils.isNotBlank(bg)) {
            for (String bgAttr : bg.split("(?<=\\)|\\w|%)\\s+(?=\\w)")) {
                if ((bgColor = PoiCssUtils.processColor(bgAttr)) != null) {
                    mapRtn.put(BACKGROUND_COLOR, bgColor);
                    break;
                }
            }
        }
        bg = style.get(BACKGROUND_COLOR);
        if (StringUtils.isNotBlank(bg) && (bgColor = PoiCssUtils.processColor(bg)) != null) {
            mapRtn.put(BACKGROUND_COLOR, bgColor);

        }
        if (bgColor != null) {
            bgColor = mapRtn.get(BACKGROUND_COLOR);
            if (!"#ffffff".equals(bgColor)) {
                return mapRtn.get(BACKGROUND_COLOR);
            }
        }
        return null;
    }

}
