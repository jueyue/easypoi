package cn.afterturn.easypoi.util;

import java.util.Map;
import java.awt.Color;
import org.slf4j.Logger;
import java.util.HashMap;
import java.util.regex.Pattern;
import java.util.regex.Matcher;
import org.slf4j.LoggerFactory;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * @version 2.0.1
 * @author Shaun Chyxion
 * @author JueYue
 */
public class PoiCssUtils {
    private static final Logger           log                       = LoggerFactory
        .getLogger(PoiCssUtils.class);
    // matches #rgb
    private static final String           COLOR_PATTERN_VALUE_SHORT = "^(#(?:[a-f]|\\d){3})$";
    // matches #rrggbb
    private static final String           COLOR_PATTERN_VALUE_LONG  = "^(#(?:[a-f]|\\d{2}){3})$";
    // matches #rgb(r, g, b)
    private static final String           COLOR_PATTERN_RGB         = "^(rgb\\s*\\(\\s*(.+)\\s*,\\s*(.+)\\s*,\\s*(.+)\\s*\\))$";

    private static final Pattern           COLOR_PATTERN_VALUE_SHORT_PATTERN  = Pattern.compile("([a-f]|\\d)");
    private static final Pattern           INT_PATTERN  = Pattern.compile("^(\\d+)(?:\\w+|%)?$");
    private static final Pattern           INT_AND_PER_PATTERN  = Pattern.compile("^(\\d*\\.?\\d+)\\s*(%)?$");

    // color name -> POI Color
    private static Map<String, HSSFColor> colors                    = new HashMap<String, HSSFColor>();
    // static init
    static {
        for (Map.Entry<Integer, HSSFColor> color : HSSFColor.getIndexHash().entrySet()) {
            colors.put(colorName(color.getValue().getClass()), color.getValue());
        }
        // light gray
        HSSFColor color = colors.get(colorName(HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getColor().getClass()));
        colors.put("lightgray", color);
        colors.put("lightgrey", color);
        // silver
        colors.put("silver", colors.get(colorName(HSSFColor.HSSFColorPredefined.GREY_40_PERCENT.getColor().getClass())));
        // darkgray
        color = colors.get(colorName(HSSFColor.HSSFColorPredefined.GREY_50_PERCENT.getColor().getClass()));
        colors.put("darkgray", color);
        colors.put("darkgrey", color);
        // gray
        color = colors.get(colorName(HSSFColor.HSSFColorPredefined.GREY_80_PERCENT.getColor().getClass()));
        colors.put("gray", color);
        colors.put("grey", color);
    }

    /**
     * get color name
     * @param color HSSFColor
     * @return color name
     */
    private static String colorName(Class<? extends HSSFColor> color) {
        return color.getSimpleName().replace("_", "").toLowerCase();
    }

    /**
     * get int value of string
     * @param strValue string value
     * @return int value
     */
    public static int getInt(String strValue) {
        int value = 0;
        if (StringUtils.isNotBlank(strValue)) {
            Matcher m = INT_PATTERN.matcher(strValue);
            if (m.find()) {
                value = Integer.parseInt(m.group(1));
            }
        }
        return value;
    }

    /**
     * check number string 
     * @param strValue string
     * @return true if string is number
     */
    public static boolean isNum(String strValue) {
        return StringUtils.isNotBlank(strValue) && strValue.matches("^\\d+(\\w+|%)?$");
    }

    /**
     * process color
     * @param color color to process
     * @return color after process
     */
    public static String processColor(String color) {
        log.info("Process Color [{}].", color);
        String colorRtn = null;
        if (StringUtils.isNotBlank(color)) {
            HSSFColor poiColor = null;
            // #rgb -> #rrggbb
            if (color.matches(COLOR_PATTERN_VALUE_SHORT)) {
                log.debug("Short Hex Color [{}] Found.", color);
                StringBuffer sbColor = new StringBuffer();
                Matcher m = COLOR_PATTERN_VALUE_SHORT_PATTERN.matcher(color);
                while (m.find()) {
                    m.appendReplacement(sbColor, "$1$1");
                }
                colorRtn = sbColor.toString();
                log.debug("Translate Short Hex Color [{}] To [{}].", color, colorRtn);
            }
            // #rrggbb
            else if (color.matches(COLOR_PATTERN_VALUE_LONG)) {
                colorRtn = color;
                log.debug("Hex Color [{}] Found, Return.", color);
            }
            // rgb(r, g, b)
            else if (color.matches(COLOR_PATTERN_RGB)) {
                Matcher m = Pattern.compile(COLOR_PATTERN_RGB).matcher(color);
                if (m.matches()) {
                    log.debug("RGB Color [{}] Found.", color);
                    colorRtn = convertColor(calcColorValue(m.group(2)), calcColorValue(m.group(3)),
                        calcColorValue(m.group(4)));
                    log.debug("Translate RGB Color [{}] To Hex [{}].", color, colorRtn);
                }
            }
            // color name, red, green, ...
            else if ((poiColor = getColor(color)) != null) {
                log.debug("Color Name [{}] Found.", color);
                short[] t = poiColor.getTriplet();
                colorRtn = convertColor(t[0], t[1], t[2]);
                log.debug("Translate Color Name [{}] To Hex [{}].", color, colorRtn);
            }
        }
        return colorRtn;
    }

    /**
     * parse color
     * @param workBook work book
     * @param color string color
     * @return HSSFColor 
     */
    public static HSSFColor parseColor(HSSFWorkbook workBook, String color) {
        HSSFColor poiColor = null;
        if (StringUtils.isNotBlank(color)) {
            Color awtColor = Color.decode(color);
            if (awtColor != null) {
                int r = awtColor.getRed();
                int g = awtColor.getGreen();
                int b = awtColor.getBlue();
                HSSFPalette palette = workBook.getCustomPalette();
                poiColor = palette.findColor((byte) r, (byte) g, (byte) b);
                if (poiColor == null) {
                    poiColor = palette.findSimilarColor(r, g, b);
                }
            }
        }
        return poiColor;
    }

    public static XSSFColor parseColor(String color) {
        XSSFColor poiColor = null;
        if (StringUtils.isNotBlank(color)) {
            Color awtColor = Color.decode(color);
            if (awtColor != null) {
                poiColor = new XSSFColor(awtColor,null);
            }
        }
        return poiColor;
    }

    // --
    // private methods

    private static HSSFColor getColor(String color) {
        return colors.get(color.replace("_", ""));
    }

    private static String convertColor(int r, int g, int b) {
        return String.format("#%02x%02x%02x", r, g, b);
    }

    public static int calcColorValue(String color) {
        int rtn = 0;
        // matches 64 or 64%
        Matcher m = INT_AND_PER_PATTERN.matcher(color);
        if (m.matches()) {
            // % not found
            if (m.group(2) == null) {
                rtn = Math.round(Float.parseFloat(m.group(1))) % 256;
            } else {
                rtn = Math.round(Float.parseFloat(m.group(1)) * 255 / 100) % 256;
            }
        }
        return rtn;
    }

}
