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

import io.github.autoffice.luckysheet.model.cell.CellData;
import io.github.autoffice.luckysheet.model.cell.CellValue;
import io.github.autoffice.luckysheet.model.cell.Sparkline;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTExtension;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTExtensionList;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.IOException;
import java.util.ArrayList;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 测试 {@link SparklineMapper} 的双向映射逻辑, 覆盖各种迷你图配置和边界情况.
 */
class SparklineMapperTest {

    private static final String EXT_URI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
    private static final String NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private static final String NS_XM = "http://schemas.microsoft.com/office/excel/2006/main";

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

    // ========== mapToLuckySheet: edge cases ==========

    @Test
    void mapToLuckySheet_nullWorksheet_doesNothing() {
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        // 模拟 null worksheet (通过不设置 extLst)
        SparklineMapper.mapToLuckySheet(sheet, luckySheet);
        // 不应该抛出异常，celldata 应该为空或保持原样
        assertNotNull(luckySheet.getCelldata());
    }

    @Test
    void mapToLuckySheet_noExtLst_doesNothing() {
        // 默认情况下 sheet 没有 extLst
        SparklineMapper.mapToLuckySheet(sheet, luckySheet);
        assertTrue(luckySheet.getCelldata().isEmpty());
    }

    @Test
    void mapToLuckySheet_extLstWithoutMatchingUri_doesNothing() throws Exception {
        // 添加一个不匹配的扩展
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTExtensionList extLst = ctWorksheet.addNewExtLst();
        CTExtension ext = extLst.addNewExt();
        ext.setUri("{DIFFERENT-URI}");

        SparklineMapper.mapToLuckySheet(sheet, luckySheet);
        assertTrue(luckySheet.getCelldata().isEmpty());
    }

    @Test
    void mapToLuckySheet_noSparklineGroups_doesNothing() throws Exception {
        // 添加正确的扩展但没有 sparklineGroup
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTExtensionList extLst = ctWorksheet.addNewExtLst();
        CTExtension ext = extLst.addNewExt();
        ext.setUri(EXT_URI);

        SparklineMapper.mapToLuckySheet(sheet, luckySheet);
        assertTrue(luckySheet.getCelldata().isEmpty());
    }

    @Test
    void mapToLuckySheet_sparklineWithoutTarget_skipped() throws Exception {
        // 创建一个没有 sqref 的 sparkline
        addSparklineToSheet("A1:A5", null, "line");

        SparklineMapper.mapToLuckySheet(sheet, luckySheet);
        assertTrue(luckySheet.getCelldata().isEmpty());
    }

    @Test
    void mapToLuckySheet_sparklineWithTypeAttribute() throws Exception {
        // 创建带有 type 属性的 sparkline
        addSparklineToSheet("A1:A5", "B2", "column");

        SparklineMapper.mapToLuckySheet(sheet, luckySheet);

        assertEquals(1, luckySheet.getCelldata().size());
        CellData cell = luckySheet.getCelldata().get(0);
        assertEquals(1, cell.getR());
        assertEquals(1, cell.getC());
        assertNotNull(cell.getV());
        assertNotNull(cell.getV().getSpl());
        assertEquals("column", cell.getV().getSpl().getType());
        assertEquals("A1:A5", cell.getV().getSpl().getDataRange());
    }

    @Test
    void mapToLuckySheet_sparklineWithoutTypeAttribute_defaultsToLine() throws Exception {
        // 创建没有 type 属性的 sparkline (默认为 line)
        addSparklineToSheet("C1:C10", "D5", null);

        SparklineMapper.mapToLuckySheet(sheet, luckySheet);

        assertEquals(1, luckySheet.getCelldata().size());
        CellData cell = luckySheet.getCelldata().get(0);
        assertEquals(4, cell.getR());
        assertEquals(3, cell.getC());
        assertEquals("line", cell.getV().getSpl().getType());
    }

    @Test
    void mapToLuckySheet_appliesSparklineToExistingCell() throws Exception {
        // 预先添加一个单元格
        CellData existingCell = new CellData();
        existingCell.setR(2);
        existingCell.setC(3);
        existingCell.setV(new CellValue());
        existingCell.getV().setV("100");
        luckySheet.getCelldata().add(existingCell);

        // 添加 sparkline 到同一位置
        addSparklineToSheet("A1:A5", "D3", "line");

        SparklineMapper.mapToLuckySheet(sheet, luckySheet);

        // 应该只有一个单元格，sparkline 被添加到现有单元格
        assertEquals(1, luckySheet.getCelldata().size());
        CellData cell = luckySheet.getCelldata().get(0);
        assertEquals(2, cell.getR());
        assertEquals(3, cell.getC());
        assertEquals("100", cell.getV().getV());
        assertNotNull(cell.getV().getSpl());
        assertEquals("line", cell.getV().getSpl().getType());
    }

    @Test
    void mapToLuckySheet_createsNewCellWhenNotExists() throws Exception {
        // 添加 sparkline 到不存在的单元格
        addSparklineToSheet("B1:B10", "E7", "column");

        SparklineMapper.mapToLuckySheet(sheet, luckySheet);

        assertEquals(1, luckySheet.getCelldata().size());
        CellData cell = luckySheet.getCelldata().get(0);
        assertEquals(6, cell.getR());
        assertEquals(4, cell.getC());
        assertNotNull(cell.getV());
        assertNotNull(cell.getV().getSpl());
    }

    @Test
    void mapToLuckySheet_cellWithNullValue_createsNewValue() throws Exception {
        // 预先添加一个 v 为 null 的单元格
        CellData existingCell = new CellData();
        existingCell.setR(1);
        existingCell.setC(1);
        existingCell.setV(null);
        luckySheet.getCelldata().add(existingCell);

        addSparklineToSheet("A1:A3", "B2", "line");

        SparklineMapper.mapToLuckySheet(sheet, luckySheet);

        assertEquals(1, luckySheet.getCelldata().size());
        CellData cell = luckySheet.getCelldata().get(0);
        assertNotNull(cell.getV());
        assertNotNull(cell.getV().getSpl());
    }

    // ========== mapToExcel: edge cases ==========

    @Test
    void mapToExcel_nullLuckySheet_doesNothing() {
        SparklineMapper.mapToExcel(null, sheet);
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        // 不应该创建 extLst
        assertTrue(!ctWorksheet.isSetExtLst());
    }

    @Test
    void mapToExcel_emptyCelldata_doesNothing() {
        luckySheet.setCelldata(new ArrayList<>());
        SparklineMapper.mapToExcel(luckySheet, sheet);
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        assertTrue(!ctWorksheet.isSetExtLst());
    }

    @Test
    void mapToExcel_cellDataWithNullV_skipped() {
        CellData cell = new CellData();
        cell.setR(0);
        cell.setC(0);
        cell.setV(null);
        luckySheet.getCelldata().add(cell);

        SparklineMapper.mapToExcel(luckySheet, sheet);
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        assertTrue(!ctWorksheet.isSetExtLst());
    }

    @Test
    void mapToExcel_cellValueWithNullSpl_skipped() {
        CellData cell = new CellData();
        cell.setR(0);
        cell.setC(0);
        cell.setV(new CellValue());
        cell.getV().setSpl(null);
        luckySheet.getCelldata().add(cell);

        SparklineMapper.mapToExcel(luckySheet, sheet);
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        assertTrue(!ctWorksheet.isSetExtLst());
    }

    @Test
    void mapToExcel_splWithNullDataRange_skipped() {
        CellData cell = new CellData();
        cell.setR(0);
        cell.setC(0);
        cell.setV(new CellValue());
        Sparkline spl = new Sparkline();
        spl.setDataRange(null);
        spl.setType("line");
        cell.getV().setSpl(spl);
        luckySheet.getCelldata().add(cell);

        SparklineMapper.mapToExcel(luckySheet, sheet);
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        assertTrue(!ctWorksheet.isSetExtLst());
    }

    @Test
    void mapToExcel_createsExtLstWhenNotExists() {
        CellData cell = createCellWithSparkline(0, 0, "A1:A5", "line");
        luckySheet.getCelldata().add(cell);

        SparklineMapper.mapToExcel(luckySheet, sheet);

        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        assertTrue(ctWorksheet.isSetExtLst());
        assertNotNull(findExtension(ctWorksheet.getExtLst()));
    }

    @Test
    void mapToExcel_reusesExistingExtLst() {
        // 先创建一个 extLst
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        ctWorksheet.addNewExtLst();

        CellData cell = createCellWithSparkline(0, 0, "A1:A5", "line");
        luckySheet.getCelldata().add(cell);

        SparklineMapper.mapToExcel(luckySheet, sheet);

        assertTrue(ctWorksheet.isSetExtLst());
        // 应该只有一个 extLst
        assertNotNull(ctWorksheet.getExtLst());
    }

    @Test
    void mapToExcel_typeNullDefaultsToLine() {
        CellData cell = createCellWithSparkline(0, 0, "A1:A5", null);
        luckySheet.getCelldata().add(cell);

        SparklineMapper.mapToExcel(luckySheet, sheet);

        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        Element extElement = (Element) findExtension(ctWorksheet.getExtLst()).getDomNode();
        NodeList groups = extElement.getElementsByTagNameNS(NS_X14, "sparklineGroup");
        assertEquals(1, groups.getLength());
        Element group = (Element) groups.item(0);
        // type="line" 时不设置 type 属性
        assertTrue(group.getAttribute("type").isEmpty());
    }

    @Test
    void mapToExcel_typeColumnSetsAttribute() {
        CellData cell = createCellWithSparkline(0, 0, "A1:A5", "column");
        luckySheet.getCelldata().add(cell);

        SparklineMapper.mapToExcel(luckySheet, sheet);

        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        Element extElement = (Element) findExtension(ctWorksheet.getExtLst()).getDomNode();
        NodeList groups = extElement.getElementsByTagNameNS(NS_X14, "sparklineGroup");
        assertEquals(1, groups.getLength());
        Element group = (Element) groups.item(0);
        assertEquals("column", group.getAttribute("type"));
    }

    @Test
    void mapToExcel_multipleSparklinesWithSameType_groupedTogether() {
        luckySheet.getCelldata().add(createCellWithSparkline(0, 0, "A1:A5", "line"));
        luckySheet.getCelldata().add(createCellWithSparkline(1, 0, "B1:B5", "line"));
        luckySheet.getCelldata().add(createCellWithSparkline(2, 0, "C1:C5", "line"));

        SparklineMapper.mapToExcel(luckySheet, sheet);

        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        Element extElement = (Element) findExtension(ctWorksheet.getExtLst()).getDomNode();
        NodeList groups = extElement.getElementsByTagNameNS(NS_X14, "sparklineGroup");
        // 应该只有一个 group
        assertEquals(1, groups.getLength());
        Element group = (Element) groups.item(0);
        NodeList sparklines = group.getElementsByTagNameNS(NS_X14, "sparkline");
        // 包含 3 个 sparkline
        assertEquals(3, sparklines.getLength());
    }

    @Test
    void mapToExcel_multipleSparklinesWithDifferentTypes_separateGroups() {
        luckySheet.getCelldata().add(createCellWithSparkline(0, 0, "A1:A5", "line"));
        luckySheet.getCelldata().add(createCellWithSparkline(1, 0, "B1:B5", "column"));
        luckySheet.getCelldata().add(createCellWithSparkline(2, 0, "C1:C5", "winloss"));

        SparklineMapper.mapToExcel(luckySheet, sheet);

        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        Element extElement = (Element) findExtension(ctWorksheet.getExtLst()).getDomNode();
        NodeList groups = extElement.getElementsByTagNameNS(NS_X14, "sparklineGroup");
        // 应该有 3 个不同的 group
        assertEquals(3, groups.getLength());
    }

    @Test
    void mapToExcel_cleansUpOldSparklineGroups() {
        // 第一次调用
        luckySheet.getCelldata().add(createCellWithSparkline(0, 0, "A1:A5", "line"));
        SparklineMapper.mapToExcel(luckySheet, sheet);

        // 第二次调用应该清理旧的 groups
        luckySheet.getCelldata().clear();
        luckySheet.getCelldata().add(createCellWithSparkline(1, 1, "B1:B5", "column"));
        SparklineMapper.mapToExcel(luckySheet, sheet);

        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        Element extElement = (Element) findExtension(ctWorksheet.getExtLst()).getDomNode();
        NodeList groups = extElement.getElementsByTagNameNS(NS_X14, "sparklineGroup");
        // 应该只有新的 group
        assertEquals(1, groups.getLength());
        Element group = (Element) groups.item(0);
        assertEquals("column", group.getAttribute("type"));
    }

    @Test
    void mapToExcel_existingNonMatchingExtension_createsNew() {
        // 添加一个不匹配的扩展
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTExtensionList extLst = ctWorksheet.addNewExtLst();
        CTExtension ext = extLst.addNewExt();
        ext.setUri("{DIFFERENT-URI}");

        luckySheet.getCelldata().add(createCellWithSparkline(0, 0, "A1:A5", "line"));
        SparklineMapper.mapToExcel(luckySheet, sheet);

        // 应该有两个扩展
        assertEquals(2, extLst.getExtList().size());
        assertNotNull(findExtension(extLst));
    }

    // ========== Round-trip tests ==========

    @Test
    void roundTrip_lineSparkline() throws Exception {
        // Excel -> Luckysheet
        addSparklineToSheet("A1:A10", "B5", "line");
        SparklineMapper.mapToLuckySheet(sheet, luckySheet);

        // Luckysheet -> Excel
        XSSFSheet sheet2 = workbook.createSheet("Test2");
        SparklineMapper.mapToExcel(luckySheet, sheet2);

        // 验证
        LuckySheet result = new LuckySheet();
        SparklineMapper.mapToLuckySheet(sheet2, result);

        assertEquals(1, result.getCelldata().size());
        CellData cell = result.getCelldata().get(0);
        assertEquals(4, cell.getR());
        assertEquals(1, cell.getC());
        assertEquals("line", cell.getV().getSpl().getType());
        assertEquals("A1:A10", cell.getV().getSpl().getDataRange());
    }

    @Test
    void roundTrip_columnSparkline() throws Exception {
        addSparklineToSheet("C1:C20", "D10", "column");
        SparklineMapper.mapToLuckySheet(sheet, luckySheet);

        XSSFSheet sheet2 = workbook.createSheet("Test2");
        SparklineMapper.mapToExcel(luckySheet, sheet2);

        LuckySheet result = new LuckySheet();
        SparklineMapper.mapToLuckySheet(sheet2, result);

        assertEquals(1, result.getCelldata().size());
        assertEquals("column", result.getCelldata().get(0).getV().getSpl().getType());
    }

    @Test
    void roundTrip_multipleSparklines() throws Exception {
        addSparklineToSheet("A1:A5", "B2", "line");
        addSparklineToSheet("C1:C5", "D3", "column");
        SparklineMapper.mapToLuckySheet(sheet, luckySheet);

        XSSFSheet sheet2 = workbook.createSheet("Test2");
        SparklineMapper.mapToExcel(luckySheet, sheet2);

        LuckySheet result = new LuckySheet();
        SparklineMapper.mapToLuckySheet(sheet2, result);

        assertEquals(2, result.getCelldata().size());
    }

    // ========== Helper methods ==========

    /**
     * 向 sheet 添加一个 sparkline 扩展
     */
    private void addSparklineToSheet(String dataRange, String target, String type) throws Exception {
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTExtensionList extLst = ctWorksheet.isSetExtLst()
                ? ctWorksheet.getExtLst()
                : ctWorksheet.addNewExtLst();

        CTExtension ext = findExtension(extLst);
        if (ext == null) {
            ext = extLst.addNewExt();
            ext.setUri(EXT_URI);
            Element extElement = (Element) ext.getDomNode();
            extElement.setAttribute("xmlns:x14", NS_X14);
        }

        Element extElement = (Element) ext.getDomNode();

        // 查找或创建 sparklineGroups
        NodeList groupsList = extElement.getElementsByTagNameNS(NS_X14, "sparklineGroups");
        Element groups;
        if (groupsList.getLength() == 0) {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = extElement.getOwnerDocument();

            groups = doc.createElementNS(NS_X14, "x14:sparklineGroups");
            groups.setAttribute("xmlns:xm", NS_XM);
            extElement.appendChild(groups);
        } else {
            groups = (Element) groupsList.item(0);
        }

        // 创建 sparklineGroup
        Document doc = extElement.getOwnerDocument();
        Element group = doc.createElementNS(NS_X14, "x14:sparklineGroup");
        if (type != null && !type.isEmpty() && !"line".equals(type)) {
            group.setAttribute("type", type);
        }

        Element sparklines = doc.createElementNS(NS_X14, "x14:sparklines");
        group.appendChild(sparklines);

        Element sparkline = doc.createElementNS(NS_X14, "x14:sparkline");

        Element f = doc.createElementNS(NS_XM, "xm:f");
        f.appendChild(doc.createTextNode(dataRange));
        sparkline.appendChild(f);

        if (target != null) {
            Element sqref = doc.createElementNS(NS_XM, "xm:sqref");
            sqref.appendChild(doc.createTextNode(target));
            sparkline.appendChild(sqref);
        }

        sparklines.appendChild(sparkline);
        groups.appendChild(group);
    }

    /**
     * 创建一个包含 sparkline 的 CellData
     */
    private CellData createCellWithSparkline(int row, int col, String dataRange, String type) {
        CellData cell = new CellData();
        cell.setR(row);
        cell.setC(col);
        cell.setV(new CellValue());
        Sparkline spl = new Sparkline();
        spl.setDataRange(dataRange);
        spl.setType(type);
        cell.getV().setSpl(spl);
        return cell;
    }

    /**
     * 查找匹配 EXT_URI 的扩展
     */
    private CTExtension findExtension(CTExtensionList extLst) {
        if (extLst == null) {
            return null;
        }
        for (CTExtension ext : extLst.getExtList()) {
            if (EXT_URI.equalsIgnoreCase(ext.getUri())) {
                return ext;
            }
        }
        return null;
    }
}
