package com.helloaldis.autoffice.luckysheet.util;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.util.Units;


public class NumberUtil {
    public static byte[] colorStringToRgb(String str) {
        byte[] rgb = new byte[]{0, 0, 0};
        if (StringUtils.isEmpty(str)) {
            return rgb;
        }

        if (str.contains("#")) {
            str = str.replace("#", "");
            str = str.trim();
            if (str.length() == 6) {

                rgb[0] = (byte) Short.parseShort(str.substring(0, 2), 16);
                rgb[1] = (byte) Short.parseShort(str.substring(2, 4), 16);
                rgb[2] = (byte) Short.parseShort(str.substring(4, 6), 16);
            }
        } else if (str.contains("rgb")) {
            str = str.replace("rgb(", "");
            str = str.replace(")", "");
            String[] split = str.split(",");
            if (split.length == 3) {
                rgb[0] = (byte) Short.parseShort(StringUtils.trim(split[0]));
                rgb[1] = (byte) Short.parseShort(StringUtils.trim(split[1]));
                rgb[2] = (byte) Short.parseShort(StringUtils.trim(split[2]));
            }
        }

        return rgb;
    }

    public static String rgbToColorString(byte[] rgb) {
        if (rgb == null || rgb.length != 3) {
            return null;
        }

        return String.format("rgb(%d,%d,%d)", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }

    public static int pixel2Twips(int pixel) {
        return (int) Math.round(Units.pixelToPoints(pixel) * Font.TWIPS_PER_POINT);
    }

    public static int twips2Pixel(int twips) {
        return Units.pointsToPixel((double) twips / Font.TWIPS_PER_POINT);
    }

    public static int pixel2PoiColWidth(int pixel) {
        return Math.round(pixel / Units.DEFAULT_CHARACTER_WIDTH * 256);
    }

    public static int pixel2CharacterLen(int pixel) {
        return Math.round(pixel / Units.DEFAULT_CHARACTER_WIDTH);
    }

    public static int characterLen2Pixel(int characterLen) {
        return Math.round(characterLen * Units.DEFAULT_CHARACTER_WIDTH);
    }

    public static int emu2Pixel(double emu) {
        return (int) Math.round(emu / Units.EMU_PER_PIXEL);
    }

    public static int pixelToEMU(float pixels) {
        return Math.round(pixels * Units.EMU_PER_PIXEL);
    }
}
