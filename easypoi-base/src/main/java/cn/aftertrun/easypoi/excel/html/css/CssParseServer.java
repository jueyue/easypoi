package cn.aftertrun.easypoi.excel.html.css;

import static cn.aftertrun.easypoi.excel.html.entity.HtmlCssConstant.*;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import cn.aftertrun.easypoi.excel.html.entity.style.CellStyleBorderEntity;
import cn.aftertrun.easypoi.excel.html.entity.style.CellStyleEntity;
import cn.aftertrun.easypoi.excel.html.entity.style.CssStyleFontEnity;
import cn.aftertrun.easypoi.util.PoiCssUtils;

/**
 * 把Css样式解析成对应的Model
 * @author JueYue
 * 2017年3月19日
 */
public class CssParseServer {

    private static final Logger      log           = LoggerFactory.getLogger(CssParseServer.class);

    @SuppressWarnings("serial")
    private final static Set<String> BORDER_STYLES = new HashSet<String>() {
                                                       {
                                                           // Specifies no border   
                                                           add(NONE);
                                                           // The same as "none", except in border conflict resolution for table elements
                                                           add(HIDDEN);
                                                           // Specifies a dotted border     
                                                           add(DOTTED);
                                                           // Specifies a dashed border     
                                                           add(DASHED);
                                                           // Specifies a solid border  
                                                           add(SOLID);
                                                           // Specifies a double border     
                                                           add(DOUBLE);
                                                       }
                                                   };

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

        parseBackground(mapStyle);

        parseBorder(mapStyle);

        return mapToCellStyleEntity(mapStyle);
    }

    private CellStyleEntity mapToCellStyleEntity(Map<String, String> mapStyle) {
        CellStyleEntity entity = new CellStyleEntity();
        entity.setAlign(mapStyle.get(TEXT_ALIGN));
        entity.setVetical(mapStyle.get(VETICAL_ALIGN));
        entity.setBackground(parseBackground(mapStyle));
        entity.setHeight(mapStyle.get(HEIGHT));
        entity.setWidth(mapStyle.get(WIDTH));
        entity.setFont(getCssStyleFontEnity(mapStyle));
        entity.setBackground(mapStyle.get(BACKGROUND_COLOR));
        entity.setBorder(getCssStyleBorderEntity(mapStyle));
        return entity;
    }

    private CellStyleBorderEntity getCssStyleBorderEntity(Map<String, String> mapStyle) {
        CellStyleBorderEntity border = new CellStyleBorderEntity();
        border.setBorderTopColor(mapStyle.get(BORDER + "-" + TOP + "-" + COLOR));
        border.setBorderBottomColor(mapStyle.get(BORDER + "-" + BOTTOM + "-" + COLOR));
        border.setBorderLeftColor(mapStyle.get(BORDER + "-" + LEFT + "-" + COLOR));
        border.setBorderRightColor(mapStyle.get(BORDER + "-" + RIGHT + "-" + COLOR));
        
        border.setBorderTopWidth(mapStyle.get(BORDER + "-" + TOP + "-" + WIDTH));
        border.setBorderBottomWidth(mapStyle.get(BORDER + "-" + BOTTOM + "-" + WIDTH));
        border.setBorderLeftWidth(mapStyle.get(BORDER + "-" + LEFT + "-" + WIDTH));
        border.setBorderRightWidth(mapStyle.get(BORDER + "-" + RIGHT + "-" + WIDTH));
        
        border.setBorderTopStyle(mapStyle.get(BORDER + "-" + TOP + "-" + STYLE));
        border.setBorderBottomStyle(mapStyle.get(BORDER + "-" + BOTTOM + "-" + STYLE));
        border.setBorderLeftStyle(mapStyle.get(BORDER + "-" + LEFT + "-" + STYLE));
        border.setBorderRightStyle(mapStyle.get(BORDER + "-" + RIGHT + "-" + STYLE));
        return border;
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
        font.setColor(style.get(COLOR));
        return font;
    }

    public void parseBorder(Map<String, String> style) {
        for (String pos : new String[] { null, TOP, RIGHT, BOTTOM, LEFT }) {
            // border[-attr]
            if (pos == null) {
                setBorderAttr(style, pos, style.get(BORDER));
                setBorderAttr(style, pos, style.get(BORDER + "-" + COLOR));
                setBorderAttr(style, pos, style.get(BORDER + "-" + WIDTH));
                setBorderAttr(style, pos, style.get(BORDER + "-" + STYLE));
            }
            // border-pos[-attr]
            else {
                setBorderAttr(style, pos, style.get(BORDER + "-" + pos));
                for (String attr : new String[] { COLOR, WIDTH, STYLE }) {
                    String attrName = BORDER + "-" + pos + "-" + attr;
                    String attrValue = style.get(attrName);
                    if (StringUtils.isNotBlank(attrValue)) {
                        style.put(attrName, attrValue);
                    }
                }
            }
        }
    }

    private void setBorderAttr(Map<String, String> mapBorder, String pos, String value) {
        if (StringUtils.isNotBlank(value)) {
            String borderColor = null;
            for (String borderAttr : value.split("\\s+")) {
                if ((borderColor = PoiCssUtils.processColor(borderAttr)) != null) {
                    setBorderAttr(mapBorder, pos, COLOR, borderColor);
                } else if (PoiCssUtils.isNum(borderAttr)) {
                    setBorderAttr(mapBorder, pos, WIDTH, borderAttr);
                } else if (BORDER_STYLES.contains(borderAttr)) {
                    setBorderAttr(mapBorder, pos, STYLE, borderAttr);
                } else {
                    log.info("Border Attr [{}] Is Not Suppoted.", borderAttr);
                }
            }
        }
    }

    private void setBorderAttr(Map<String, String> mapBorder, String pos, String attr,
                               String value) {
        if (StringUtils.isNotBlank(pos)) {
            mapBorder.put(BORDER + "-" + pos + "-" + attr, value);
        } else {
            for (String name : new String[] { TOP, RIGHT, BOTTOM, LEFT }) {
                mapBorder.put(BORDER + "-" + name + "-" + attr, value);
            }
        }
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

    private String parseBackground(Map<String, String> style) {
        String bg = style.get(BACKGROUND);
        String bgColor = null;
        if (StringUtils.isNotBlank(bg)) {
            for (String bgAttr : bg.split("(?<=\\)|\\w|%)\\s+(?=\\w)")) {
                if ((bgColor = PoiCssUtils.processColor(bgAttr)) != null) {
                    style.put(BACKGROUND_COLOR, bgColor);
                    break;
                }
            }
        }
        bg = style.get(BACKGROUND_COLOR);
        if (StringUtils.isNotBlank(bg) && (bgColor = PoiCssUtils.processColor(bg)) != null) {
            style.put(BACKGROUND_COLOR, bgColor);

        }
        if (bgColor != null) {
            bgColor = style.get(BACKGROUND_COLOR);
            if ("#ffffff".equals(bgColor)) {
                style.remove(BACKGROUND_COLOR);
            }
        }
        return null;
    }

}
