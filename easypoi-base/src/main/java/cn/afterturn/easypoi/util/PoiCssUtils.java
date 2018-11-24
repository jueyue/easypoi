package cn.afterturn.easypoi.util;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author Shaun Chyxion
 * @author JueYue
 * @version 2.0.1
 */
public class PoiCssUtils {
    private static final Logger log = LoggerFactory
            .getLogger(PoiCssUtils.class);
    /**
     * matches #rgb
     */
    private static final String COLOR_PATTERN_VALUE_SHORT = "^(#(?:[a-f]|\\d){3})$";
    /**
     * matches #rrggbb
     */
    private static final String COLOR_PATTERN_VALUE_LONG = "^(#(?:[a-f]|\\d{2}){3})$";
    /**
     * matches #rgb(r, g, b)
     **/
    private static final String COLOR_PATTERN_RGB = "^(rgb\\s*\\(\\s*(.+)\\s*,\\s*(.+)\\s*,\\s*(.+)\\s*\\))$";

    private static final Pattern COLOR_PATTERN_VALUE_SHORT_PATTERN = Pattern.compile("([a-f]|\\d)");
    private static final Pattern INT_PATTERN = Pattern.compile("^(\\d+)(?:\\w+|%)?$");
    private static final Pattern INT_AND_PER_PATTERN = Pattern.compile("^(\\d*\\.?\\d+)\\s*(%)?$");

    /**
     * get color name
     *
     * @param color HSSFColor
     * @return color name
     */
    private static String colorName(Class<? extends HSSFColor> color) {
        return color.getSimpleName().replace("_", "").toLowerCase();
    }

    /**
     * get int value of string
     *
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
     *
     * @param strValue string
     * @return true if string is number
     */
    public static boolean isNum(String strValue) {
        return StringUtils.isNotBlank(strValue) && strValue.matches("^\\d+(\\w+|%)?$");
    }

    /**
     * process color
     *
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
            else if ((poiColor = getHssfColor(color)) != null) {
                log.debug("Color Name [{}] Found.", color);
                short[] t = poiColor.getTriplet();
                colorRtn = convertColor(t[0], t[1], t[2]);
                log.debug("Translate Color Name [{}] To Hex [{}].", color, colorRtn);
            }
        }
        return colorRtn;
    }

    private static HSSFColor getHssfColor(String color) {
        String tmpColor = color.replace("_", "").toUpperCase();
        switch (tmpColor) {
            case "BLACK":
                return HSSFColor.HSSFColorPredefined.BLACK.getColor();
            case "BROWN":
                return HSSFColor.HSSFColorPredefined.BROWN.getColor();
            case "OLIVEGREEN":
                return HSSFColor.HSSFColorPredefined.OLIVE_GREEN.getColor();
            case "DARKGREEN":
                return HSSFColor.HSSFColorPredefined.DARK_GREEN.getColor();
            case "DARKTEAL":
                return HSSFColor.HSSFColorPredefined.DARK_TEAL.getColor();
            case "DARKBLUE":
                return HSSFColor.HSSFColorPredefined.DARK_BLUE.getColor();
            case "INDIGO":
                return HSSFColor.HSSFColorPredefined.INDIGO.getColor();
            case "GREY80PERCENT":
                return HSSFColor.HSSFColorPredefined.GREY_80_PERCENT.getColor();
            case "ORANGE":
                return HSSFColor.HSSFColorPredefined.ORANGE.getColor();
            case "DARKYELLOW":
                return HSSFColor.HSSFColorPredefined.DARK_YELLOW.getColor();
            case "GREEN":
                return HSSFColor.HSSFColorPredefined.GREEN.getColor();
            case "TEAL":
                return HSSFColor.HSSFColorPredefined.TEAL.getColor();
            case "BLUE":
                return HSSFColor.HSSFColorPredefined.BLUE.getColor();
            case "BLUEGREY":
                return HSSFColor.HSSFColorPredefined.BLUE_GREY.getColor();
            case "GREY50PERCENT":
                return HSSFColor.HSSFColorPredefined.GREY_50_PERCENT.getColor();
            case "RED":
                return HSSFColor.HSSFColorPredefined.RED.getColor();
            case "LIGHTORANGE":
                return HSSFColor.HSSFColorPredefined.LIGHT_ORANGE.getColor();
            case "LIME":
                return HSSFColor.HSSFColorPredefined.LIME.getColor();
            case "SEAGREEN":
                return HSSFColor.HSSFColorPredefined.SEA_GREEN.getColor();
            case "AQUA":
                return HSSFColor.HSSFColorPredefined.AQUA.getColor();
            case "LIGHTBLUE":
                return HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getColor();
            case "VIOLET":
                return HSSFColor.HSSFColorPredefined.VIOLET.getColor();
            case "GREY40PERCENT":
                return HSSFColor.HSSFColorPredefined.GREY_40_PERCENT.getColor();
            case "GOLD":
                return HSSFColor.HSSFColorPredefined.GOLD.getColor();
            case "YELLOW":
                return HSSFColor.HSSFColorPredefined.YELLOW.getColor();
            case "BRIGHTGREEN":
                return HSSFColor.HSSFColorPredefined.BRIGHT_GREEN.getColor();
            case "TURQUOISE":
                return HSSFColor.HSSFColorPredefined.TURQUOISE.getColor();
            case "DARKRED":
                return HSSFColor.HSSFColorPredefined.DARK_RED.getColor();
            case "SKYBLUE":
                return HSSFColor.HSSFColorPredefined.SKY_BLUE.getColor();
            case "PLUM":
                return HSSFColor.HSSFColorPredefined.PLUM.getColor();
            case "GREY25PERCENT":
                return HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getColor();
            case "ROSE":
                return HSSFColor.HSSFColorPredefined.ROSE.getColor();
            case "LIGHTYELLOW":
                return HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getColor();
            case "LIGHTGREEN":
                return HSSFColor.HSSFColorPredefined.LIGHT_GREEN.getColor();
            case "LIGHTTURQUOISE":
                return HSSFColor.HSSFColorPredefined.LIGHT_TURQUOISE.getColor();
            case "PALEBLUE":
                return HSSFColor.HSSFColorPredefined.PALE_BLUE.getColor();
            case "LAVENDER":
                return HSSFColor.HSSFColorPredefined.LAVENDER.getColor();
            case "WHITE":
                return HSSFColor.HSSFColorPredefined.WHITE.getColor();
            case "CORNFLOWERBLUE":
                return HSSFColor.HSSFColorPredefined.CORNFLOWER_BLUE.getColor();
            case "LEMONCHIFFON":
                return HSSFColor.HSSFColorPredefined.LEMON_CHIFFON.getColor();
            case "MAROON":
                return HSSFColor.HSSFColorPredefined.MAROON.getColor();
            case "ORCHID":
                return HSSFColor.HSSFColorPredefined.ORANGE.getColor();
            case "CORAL":
                return HSSFColor.HSSFColorPredefined.CORAL.getColor();
            case "ROYALBLUE":
                return HSSFColor.HSSFColorPredefined.ROYAL_BLUE.getColor();
            case "LIGHTCORNFLOWERBLUE":
                return HSSFColor.HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE.getColor();
            case "TAN":
                return HSSFColor.HSSFColorPredefined.TAN.getColor();
        }
        return null;
    }

    /**
     * parse color
     *
     * @param workBook work book
     * @param color    string color
     * @return HSSFColor
     */
    public static HSSFColor parseColor(HSSFWorkbook workBook, String color) {
        HSSFColor poiColor = null;
        if (StringUtils.isNotBlank(color)) {
            Color awtColor = getAwtColor(color);
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

    private static Color getAwtColor(String color) {
        HSSFColor hssfColor = getHssfColor(color);
        if (hssfColor != null) {
            short[] t = hssfColor.getTriplet();
            return new Color(t[0], t[1], t[2]);
        }
        String tmpColor = color.replace("_", "").toLowerCase();
        switch (tmpColor) {
            case "lightgray":
                return Color.LIGHT_GRAY;
            case "gray":
                return Color.GRAY;
            case "darkgray":
                return Color.DARK_GRAY;
            case "pink":
                return Color.PINK;
            case "magenta":
                return Color.MAGENTA;
            case "cyan":
                return Color.CYAN;
        }
        return Color.decode(color);
    }

    public static XSSFColor parseColor(String color) {
        XSSFColor poiColor = null;
        if (StringUtils.isNotBlank(color)) {
            Color awtColor = getAwtColor(color);
            if (awtColor != null) {
                poiColor = new XSSFColor(awtColor, null);
            }
        }
        return poiColor;
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
