package com.helloaldis.autoffice.luckysheet.util;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;

class NumberUtilTest {

    @Test
    void colorStringToRgb() {
        assertArrayEquals(new byte[]{51, 51, 51}, NumberUtil.colorStringToRgb("rgb(51,51,51)"));
        assertArrayEquals(new byte[]{-1, -1, -1}, NumberUtil.colorStringToRgb("rgb(255,255,255)"));
    }

    @Test
    void rgbToColorString() {
        assertEquals("rgb(51,51,51)", NumberUtil.rgbToColorString(new byte[]{51, 51, 51}));
        assertEquals("rgb(255,255,255)", NumberUtil.rgbToColorString(new byte[]{-1, -1, -1}));
    }
}
