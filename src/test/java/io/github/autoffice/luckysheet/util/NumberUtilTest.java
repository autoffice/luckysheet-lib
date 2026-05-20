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

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class NumberUtilTest {

    // --- colorStringToRgb tests ---

    @Test
    void colorStringToRgb_rgbFormat() {
        assertArrayEquals(new byte[]{51, 51, 51}, NumberUtil.colorStringToRgb("rgb(51,51,51)"));
        assertArrayEquals(new byte[]{-1, -1, -1}, NumberUtil.colorStringToRgb("rgb(255,255,255)"));
        assertArrayEquals(new byte[]{0, 0, 0}, NumberUtil.colorStringToRgb("rgb(0,0,0)"));
        // with spaces
        assertArrayEquals(new byte[]{10, 20, 30}, NumberUtil.colorStringToRgb("rgb(10, 20, 30)"));
    }

    @Test
    void colorStringToRgb_hexFormat() {
        assertArrayEquals(new byte[]{-1, 0, 0}, NumberUtil.colorStringToRgb("#ff0000"));
        assertArrayEquals(new byte[]{0, -1, 0}, NumberUtil.colorStringToRgb("#00ff00"));
        assertArrayEquals(new byte[]{0, 0, -1}, NumberUtil.colorStringToRgb("#0000ff"));
        assertArrayEquals(new byte[]{51, 51, 51}, NumberUtil.colorStringToRgb("#333333"));
    }

    @Test
    void colorStringToRgb_emptyAndNull() {
        assertArrayEquals(new byte[]{0, 0, 0}, NumberUtil.colorStringToRgb(null));
        assertArrayEquals(new byte[]{0, 0, 0}, NumberUtil.colorStringToRgb(""));
    }

    @Test
    void colorStringToRgb_invalidFormat() {
        // No rgb or # prefix - returns default
        assertArrayEquals(new byte[]{0, 0, 0}, NumberUtil.colorStringToRgb("invalid"));
        // Short hex (not 6 chars) - returns default
        assertArrayEquals(new byte[]{0, 0, 0}, NumberUtil.colorStringToRgb("#fff"));
    }

    // --- rgbToColorString tests ---

    @Test
    void rgbToColorString_validInput() {
        assertEquals("rgb(51,51,51)", NumberUtil.rgbToColorString(new byte[]{51, 51, 51}));
        assertEquals("rgb(255,255,255)", NumberUtil.rgbToColorString(new byte[]{-1, -1, -1}));
        assertEquals("rgb(0,0,0)", NumberUtil.rgbToColorString(new byte[]{0, 0, 0}));
        assertEquals("rgb(128,64,32)", NumberUtil.rgbToColorString(new byte[]{-128, 64, 32}));
    }

    @Test
    void rgbToColorString_nullAndInvalidLength() {
        assertNull(NumberUtil.rgbToColorString(null));
        assertNull(NumberUtil.rgbToColorString(new byte[]{1, 2}));
        assertNull(NumberUtil.rgbToColorString(new byte[]{1, 2, 3, 4}));
    }

    // --- Unit conversion tests ---

    @Test
    void pixel2Twips_variousValues() {
        assertEquals(0, NumberUtil.pixel2Twips(0));
        assertTrue(NumberUtil.pixel2Twips(1) > 0);
        assertTrue(NumberUtil.pixel2Twips(100) > NumberUtil.pixel2Twips(50));
    }

    @Test
    void twips2Pixel_variousValues() {
        assertEquals(0, NumberUtil.twips2Pixel(0));
        assertTrue(NumberUtil.twips2Pixel(20) >= 0);
        assertTrue(NumberUtil.twips2Pixel(1000) > NumberUtil.twips2Pixel(100));
    }

    @Test
    void pixel2PoiColWidth_variousValues() {
        assertEquals(0, NumberUtil.pixel2PoiColWidth(0));
        assertTrue(NumberUtil.pixel2PoiColWidth(10) > 0);
        assertTrue(NumberUtil.pixel2PoiColWidth(100) > NumberUtil.pixel2PoiColWidth(10));
    }

    @Test
    void pixel2CharacterLen_variousValues() {
        assertEquals(0, NumberUtil.pixel2CharacterLen(0));
        assertTrue(NumberUtil.pixel2CharacterLen(100) > 0);
    }

    @Test
    void characterLen2Pixel_variousValues() {
        assertEquals(0, NumberUtil.characterLen2Pixel(0));
        assertTrue(NumberUtil.characterLen2Pixel(10) > 0);
    }

    @Test
    void emu2Pixel_variousValues() {
        assertEquals(0, NumberUtil.emu2Pixel(0));
        assertTrue(NumberUtil.emu2Pixel(914400) > 0);
    }

    @Test
    void pixelToEMU_variousValues() {
        assertEquals(0, NumberUtil.pixelToEMU(0));
        assertTrue(NumberUtil.pixelToEMU(1.0f) > 0);
        assertTrue(NumberUtil.pixelToEMU(100.0f) > NumberUtil.pixelToEMU(10.0f));
    }

    @Test
    void roundTripConversions() {
        // pixel -> twips -> pixel should be approximately equal
        int original = 100;
        int twips = NumberUtil.pixel2Twips(original);
        int backToPixel = NumberUtil.twips2Pixel(twips);
        assertTrue(Math.abs(original - backToPixel) <= 1);
    }
}
