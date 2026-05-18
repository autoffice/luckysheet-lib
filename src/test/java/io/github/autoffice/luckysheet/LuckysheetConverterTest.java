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
package io.github.autoffice.luckysheet;

import io.github.autoffice.luckysheet.model.DefinedName;
import io.github.autoffice.luckysheet.model.LuckyFile;
import io.github.autoffice.luckysheet.model.sheet.Authority;
import io.github.autoffice.luckysheet.model.sheet.ConditionFormat;
import io.github.autoffice.luckysheet.model.sheet.DataVerification;
import io.github.autoffice.luckysheet.model.sheet.Group;
import io.github.autoffice.luckysheet.model.sheet.Hyperlink;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.model.sheet.Range;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class LuckysheetConverterTest {
    private static final String OUTPUT = "./target/";

    @Test
    void fullTesting() throws IOException, InvalidFormatException {
        URL resource = getClass().getResource("/full.json");
        assertNotNull(resource, "Resource not found");
        LuckysheetConverter.luckysheetToExcel(resource.getPath(), OUTPUT + "full.xlsx");
        LuckysheetConverter.excelToLuckySheetFile(OUTPUT + "full.xlsx", OUTPUT + "full1.json");
        LuckysheetConverter.luckysheetToExcel(OUTPUT + "full1.json", OUTPUT + "full1.xlsx");
        LuckysheetConverter.excelToLuckySheetFile(OUTPUT + "full1.xlsx", OUTPUT + "full2.json");
        LuckysheetConverter.luckysheetToExcel(OUTPUT + "full2.json", Files.newOutputStream(Paths.get(OUTPUT + "full2.xlsx")));
    }

    @Test
    void testXlsxToLuckysheet() throws IOException, InvalidFormatException {
        URL resource = getClass().getResource("/test.xlsx");
        assertNotNull(resource, "Resource not found");
        LuckysheetConverter.excelToLuckySheetFile(resource.getPath(), OUTPUT + "test.json");
        LuckysheetConverter.luckysheetToExcel(OUTPUT + "test.json", OUTPUT + "test.xlsx");
    }

    /**
     * 验证新增特性 (超链接、数据验证、条件格式、自动筛选、工作表保护、行列分组、命名范围)
     * 的 JSON → Excel → JSON round-trip.
     */
    @Test
    void newFeaturesTesting() throws IOException, InvalidFormatException {
        URL resource = getClass().getResource("/newFeatures.json");
        assertNotNull(resource, "Resource not found");
        String xlsx = OUTPUT + "newFeatures.xlsx";
        String json = OUTPUT + "newFeatures.out.json";
        LuckysheetConverter.luckysheetToExcel(resource.getPath(), xlsx);
        LuckyFile roundTrip = LuckysheetConverter.excelToLuckySheet(xlsx);

        assertNotNull(roundTrip.getSheets());
        assertFalse(roundTrip.getSheets().isEmpty());
        LuckySheet sheet0 = roundTrip.getSheets().get(0);

        // 超链接
        Map<String, Hyperlink> hyperlink = sheet0.getHyperlink();
        assertNotNull(hyperlink, "hyperlink missing");
        Hyperlink link = hyperlink.get("0_1");
        assertNotNull(link, "hyperlink 0_1 missing");
        assertEquals("https://www.google.com", link.getLinkAddress());

        // 数据验证
        Map<String, DataVerification> dv = sheet0.getDataVerification();
        assertNotNull(dv, "dataVerification missing");
        assertTrue(dv.containsKey("1_2"), "expected dropdown at 1_2");
        DataVerification dropdown = dv.get("1_2");
        assertEquals("dropdown", dropdown.getType());

        // 条件格式: >=1 条
        List<ConditionFormat> cfs = sheet0.getLuckysheet_conditionformat_save();
        assertNotNull(cfs, "conditional formatting missing");
        assertFalse(cfs.isEmpty(), "conditional formatting missing entries");

        // 自动筛选范围 round-trip (详细列条件有损, 仅验证范围)
        Range filterSelect = sheet0.getFilter_select();
        assertNotNull(filterSelect, "filter_select missing");

        // 工作表保护
        Authority authority = sheet0.getConfig().getAuthority();
        assertNotNull(authority, "authority missing");
        assertEquals(Integer.valueOf(1), authority.getSheet());

        // 行/列分组
        List<Group> rowGroups = sheet0.getRowGroup();
        assertNotNull(rowGroups, "rowGroup missing");
        assertFalse(rowGroups.isEmpty());

        List<Group> colGroups = sheet0.getColGroup();
        assertNotNull(colGroups, "colGroup missing");
        assertFalse(colGroups.isEmpty());

        // 命名范围 (workbook 级)
        List<DefinedName> definedNames = roundTrip.getDefinedNames();
        assertNotNull(definedNames, "definedNames missing");
        assertFalse(definedNames.isEmpty());
        assertEquals("MyRange", definedNames.get(0).getName());

        // 输出最终 json 以便调试
        LuckysheetConverter.excelToLuckySheetFile(xlsx, json);
    }
}
