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
package io.github.autoffice.luckysheet.mapper;

import io.github.autoffice.luckysheet.model.sheet.Hyperlink;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 测试 {@link HyperlinkMapper} 的双向映射逻辑, 覆盖各种超链接类型和边界情况.
 */
class HyperlinkMapperTest {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private LuckySheet luckySheet;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Test");
        luckySheet = new LuckySheet();
        luckySheet.setName("Test");
        luckySheet.setIndex("0");
    }

    @AfterEach
    void tearDown() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }

    // ========== mapToLuckySheet ==========

    @Test
    void mapToLuckySheet_emptyHyperlinks_doesNothing() {
        HyperlinkMapper.mapToLuckySheet(sheet, luckySheet);
        assertNull(luckySheet.getHyperlink());
    }

    @Test
    void mapToLuckySheet_urlType() {
        XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
        link.setAddress("https://example.com");
        link.setCellReference("A1");
        sheet.addHyperlink(link);

        HyperlinkMapper.mapToLuckySheet(sheet, luckySheet);

        Map<String, Hyperlink> map = luckySheet.getHyperlink();
        assertNotNull(map);
        assertEquals(1, map.size());
        Hyperlink h = map.get("0_0");
        assertNotNull(h);
        assertEquals("https://example.com", h.getLinkAddress());
        assertEquals("external", h.getLinkType());
    }

    @Test
    void mapToLuckySheet_emailType() {
        XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.EMAIL);
        link.setAddress("mailto:test@example.com");
        link.setCellReference("B2");
        sheet.addHyperlink(link);

        HyperlinkMapper.mapToLuckySheet(sheet, luckySheet);

        Map<String, Hyperlink> map = luckySheet.getHyperlink();
        assertNotNull(map);
        Hyperlink h = map.get("1_1");
        assertNotNull(h);
        assertEquals("mailto:test@example.com", h.getLinkAddress());
        assertEquals("external", h.getLinkType());
    }

    @Test
    void mapToLuckySheet_documentType() {
        XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
        link.setAddress("Sheet2!A1");
        link.setCellReference("C3");
        sheet.addHyperlink(link);

        HyperlinkMapper.mapToLuckySheet(sheet, luckySheet);

        Map<String, Hyperlink> map = luckySheet.getHyperlink();
        assertNotNull(map);
        Hyperlink h = map.get("2_2");
        assertNotNull(h);
        assertEquals("Sheet2!A1", h.getLinkAddress());
        assertEquals("internal", h.getLinkType());
    }

    @Test
    void mapToLuckySheet_withTooltip() {
        XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
        link.setAddress("https://example.com");
        link.setTooltip("Visit our website");
        link.setCellReference("A1");
        sheet.addHyperlink(link);

        HyperlinkMapper.mapToLuckySheet(sheet, luckySheet);

        Hyperlink h = luckySheet.getHyperlink().get("0_0");
        assertNotNull(h);
        assertEquals("Visit our website", h.getLinkTooltip());
    }

    @Test
    void mapToLuckySheet_withoutTooltip() {
        XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
        link.setAddress("https://example.com");
        link.setCellReference("A1");
        sheet.addHyperlink(link);

        HyperlinkMapper.mapToLuckySheet(sheet, luckySheet);

        Hyperlink h = luckySheet.getHyperlink().get("0_0");
        assertNotNull(h);
        assertNull(h.getLinkTooltip());
    }

    @Test
    void mapToLuckySheet_multipleHyperlinks() {
        XSSFHyperlink link1 = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
        link1.setAddress("https://example1.com");
        link1.setCellReference("A1");
        sheet.addHyperlink(link1);

        XSSFHyperlink link2 = workbook.getCreationHelper().createHyperlink(HyperlinkType.EMAIL);
        link2.setAddress("mailto:test@example.com");
        link2.setCellReference("B2");
        sheet.addHyperlink(link2);

        XSSFHyperlink link3 = workbook.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
        link3.setAddress("Sheet2!A1");
        link3.setCellReference("C3");
        sheet.addHyperlink(link3);

        HyperlinkMapper.mapToLuckySheet(sheet, luckySheet);

        Map<String, Hyperlink> map = luckySheet.getHyperlink();
        assertNotNull(map);
        assertEquals(3, map.size());
        assertTrue(map.containsKey("0_0"));
        assertTrue(map.containsKey("1_1"));
        assertTrue(map.containsKey("2_2"));
    }

    // ========== mapToExcel ==========

    @Test
    void mapToExcel_nullMap_doesNothing() {
        HyperlinkMapper.mapToExcel(null, sheet);
        assertEquals(0, sheet.getHyperlinkList().size());
    }

    @Test
    void mapToExcel_emptyMap_doesNothing() {
        HyperlinkMapper.mapToExcel(new HashMap<>(), sheet);
        assertEquals(0, sheet.getHyperlinkList().size());
    }

    @Test
    void mapToExcel_nullLinkAddress_skipped() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress(null);
        h.setLinkType("external");
        map.put("0_0", h);

        HyperlinkMapper.mapToExcel(map, sheet);
        assertEquals(0, sheet.getHyperlinkList().size());
    }

    @Test
    void mapToExcel_blankLinkAddress_skipped() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("   ");
        h.setLinkType("external");
        map.put("0_0", h);

        HyperlinkMapper.mapToExcel(map, sheet);
        assertEquals(0, sheet.getHyperlinkList().size());
    }

    @Test
    void mapToExcel_invalidKeyFormat_noUnderscore_skipped() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("https://example.com");
        h.setLinkType("external");
        map.put("00", h);

        HyperlinkMapper.mapToExcel(map, sheet);
        assertEquals(0, sheet.getHyperlinkList().size());
    }

    @Test
    void mapToExcel_invalidKeyFormat_underscoreAtStart_skipped() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("https://example.com");
        h.setLinkType("external");
        map.put("_0", h);

        HyperlinkMapper.mapToExcel(map, sheet);
        assertEquals(0, sheet.getHyperlinkList().size());
    }

    @Test
    void mapToExcel_invalidKeyFormat_underscoreAtEnd_skipped() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("https://example.com");
        h.setLinkType("external");
        map.put("0_", h);

        HyperlinkMapper.mapToExcel(map, sheet);
        assertEquals(0, sheet.getHyperlinkList().size());
    }

    @Test
    void mapToExcel_invalidKeyFormat_nonNumeric_skipped() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("https://example.com");
        h.setLinkType("external");
        map.put("a_b", h);

        HyperlinkMapper.mapToExcel(map, sheet);
        assertEquals(0, sheet.getHyperlinkList().size());
    }

    @Test
    void mapToExcel_nullKey_skipped() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("https://example.com");
        h.setLinkType("external");
        map.put(null, h);

        HyperlinkMapper.mapToExcel(map, sheet);
        assertEquals(0, sheet.getHyperlinkList().size());
    }

    @Test
    void mapToExcel_internalLinkType_becomesDocument() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("Sheet2!A1");
        h.setLinkType("internal");
        map.put("0_0", h);

        HyperlinkMapper.mapToExcel(map, sheet);

        List<XSSFHyperlink> links = sheet.getHyperlinkList();
        assertEquals(1, links.size());
        assertEquals(HyperlinkType.DOCUMENT, links.get(0).getType());
        assertEquals("Sheet2!A1", links.get(0).getAddress());
    }

    @Test
    void mapToExcel_externalWithMailto_becomesEmail() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("mailto:test@example.com");
        h.setLinkType("external");
        map.put("1_1", h);

        HyperlinkMapper.mapToExcel(map, sheet);

        List<XSSFHyperlink> links = sheet.getHyperlinkList();
        assertEquals(1, links.size());
        assertEquals(HyperlinkType.EMAIL, links.get(0).getType());
        assertEquals("mailto:test@example.com", links.get(0).getAddress());
    }

    @Test
    void mapToExcel_externalWithHttp_becomesUrl() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("http://example.com");
        h.setLinkType("external");
        map.put("2_2", h);

        HyperlinkMapper.mapToExcel(map, sheet);

        List<XSSFHyperlink> links = sheet.getHyperlinkList();
        assertEquals(1, links.size());
        assertEquals(HyperlinkType.URL, links.get(0).getType());
        assertEquals("http://example.com", links.get(0).getAddress());
    }

    @Test
    void mapToExcel_withTooltip() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("https://example.com");
        h.setLinkType("external");
        h.setLinkTooltip("Click here");
        map.put("0_0", h);

        HyperlinkMapper.mapToExcel(map, sheet);

        List<XSSFHyperlink> links = sheet.getHyperlinkList();
        assertEquals(1, links.size());
        assertEquals("Click here", links.get(0).getTooltip());
    }

    @Test
    void mapToExcel_withoutTooltip() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("https://example.com");
        h.setLinkType("external");
        map.put("0_0", h);

        HyperlinkMapper.mapToExcel(map, sheet);

        List<XSSFHyperlink> links = sheet.getHyperlinkList();
        assertEquals(1, links.size());
        assertNull(links.get(0).getTooltip());
    }

    @Test
    void mapToExcel_nullTooltip() {
        Map<String, Hyperlink> map = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("https://example.com");
        h.setLinkType("external");
        h.setLinkTooltip(null);
        map.put("0_0", h);

        HyperlinkMapper.mapToExcel(map, sheet);

        List<XSSFHyperlink> links = sheet.getHyperlinkList();
        assertEquals(1, links.size());
        assertNull(links.get(0).getTooltip());
    }

    // ========== Round-trip ==========

    @Test
    void roundTrip_urlHyperlink() {
        Map<String, Hyperlink> src = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("https://example.com");
        h.setLinkType("external");
        h.setLinkTooltip("Example site");
        src.put("5_10", h);

        HyperlinkMapper.mapToExcel(src, sheet);

        LuckySheet result = new LuckySheet();
        HyperlinkMapper.mapToLuckySheet(sheet, result);

        Map<String, Hyperlink> dest = result.getHyperlink();
        assertNotNull(dest);
        assertEquals(1, dest.size());
        Hyperlink resultH = dest.get("5_10");
        assertNotNull(resultH);
        assertEquals("https://example.com", resultH.getLinkAddress());
        assertEquals("external", resultH.getLinkType());
        assertEquals("Example site", resultH.getLinkTooltip());
    }

    @Test
    void roundTrip_internalHyperlink() {
        Map<String, Hyperlink> src = new HashMap<>();
        Hyperlink h = new Hyperlink();
        h.setLinkAddress("Sheet2!B5");
        h.setLinkType("internal");
        src.put("0_0", h);

        HyperlinkMapper.mapToExcel(src, sheet);

        LuckySheet result = new LuckySheet();
        HyperlinkMapper.mapToLuckySheet(sheet, result);

        Map<String, Hyperlink> dest = result.getHyperlink();
        assertNotNull(dest);
        Hyperlink resultH = dest.get("0_0");
        assertNotNull(resultH);
        assertEquals("Sheet2!B5", resultH.getLinkAddress());
        assertEquals("internal", resultH.getLinkType());
    }
}
