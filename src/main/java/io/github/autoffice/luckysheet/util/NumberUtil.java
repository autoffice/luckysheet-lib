/*
 * Copyright © 2025 AutOffice (hello.aldis@qq.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package io.github.autoffice.luckysheet.util;

import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.util.Units;

/**
 * 数值与单位转换工具类.
 *
 * <p>提供颜色格式转换以及像素、Twips、EMU、字符宽度等单位之间的互相转换.</p>
 */
@NoArgsConstructor(access = AccessLevel.PRIVATE)
public final class NumberUtil {
    /**
     * 将颜色字符串 (如 "rgb(255,0,0)" 或 "#ff0000") 转换为 RGB 字节数组.
     *
     * @param str 颜色字符串
     * @return RGB 字节数组 (长度为3)
     */
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

    /**
     * 将 RGB 字节数组转换为 "rgb(r,g,b)" 格式的颜色字符串.
     *
     * @param rgb RGB 字节数组 (长度为3)
     * @return "rgb(r,g,b)" 格式字符串, 输入无效时返回 null
     */
    public static String rgbToColorString(byte[] rgb) {
        if (rgb == null || rgb.length != 3) {
            return null;
        }

        return String.format("rgb(%d,%d,%d)", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }

    /**
     * 将像素值转换为 Twips 单位.
     *
     * @param pixel 像素值
     * @return Twips 值
     */
    public static int pixel2Twips(int pixel) {
        return (int) Math.round(Units.pixelToPoints(pixel) * Font.TWIPS_PER_POINT);
    }

    /**
     * 将 Twips 单位转换为像素值.
     *
     * @param twips Twips 值
     * @return 像素值
     */
    public static int twips2Pixel(int twips) {
        return Units.pointsToPixel((double) twips / Font.TWIPS_PER_POINT);
    }

    /**
     * 将像素值转换为 POI 列宽单位 (1/256 字符宽度).
     *
     * @param pixel 像素值
     * @return POI 列宽值
     */
    public static int pixel2PoiColWidth(int pixel) {
        return Math.round(pixel / Units.DEFAULT_CHARACTER_WIDTH * 256);
    }

    /**
     * 将像素值转换为字符宽度.
     *
     * @param pixel 像素值
     * @return 字符宽度
     */
    public static int pixel2CharacterLen(int pixel) {
        return Math.round(pixel / Units.DEFAULT_CHARACTER_WIDTH);
    }

    /**
     * 将字符宽度转换为像素值.
     *
     * @param characterLen 字符宽度
     * @return 像素值
     */
    public static int characterLen2Pixel(int characterLen) {
        return Math.round(characterLen * Units.DEFAULT_CHARACTER_WIDTH);
    }

    /**
     * 将 EMU (English Metric Units) 转换为像素值.
     *
     * @param emu EMU 值
     * @return 像素值
     */
    public static int emu2Pixel(double emu) {
        return (int) Math.round(emu / Units.EMU_PER_PIXEL);
    }

    /**
     * 将像素值转换为 EMU (English Metric Units).
     *
     * @param pixels 像素值
     * @return EMU 值
     */
    public static int pixelToEMU(float pixels) {
        return Math.round(pixels * Units.EMU_PER_PIXEL);
    }
}
