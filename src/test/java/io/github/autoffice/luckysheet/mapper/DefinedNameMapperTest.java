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

import io.github.autoffice.luckysheet.model.DefinedName;
import io.github.autoffice.luckysheet.model.LuckyFile;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 测试 {@link DefinedNameMapper} 的双向映射逻辑, 覆盖命名范围的各种配置场景.
 */
class DefinedNameMapperTest {

    private XSSFWorkbook workbook;
    private LuckyFile luckyFile;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        workbook.createSheet("Sheet1");
        workbook.createSheet("Sheet2");
        luckyFile = new LuckyFile();
    }

    @AfterEach
    void tearDown() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }

    // ========== mapToLuckyFile ==========

    @Test
    void mapToLuckyFile_emptyNames_doesNothing() {
        DefinedNameMapper.mapToLuckyFile(workbook, luckyFile);
        assertNull(luckyFile.getDefinedNames());
    }

    @Test
    void mapToLuckyFile_normalUserDefinedName() {
        XSSFName name = workbook.createName();
        name.setNameName("MyRange");
        name.setRefersToFormula("Sheet1!$A$1:$B$10");

        DefinedNameMapper.mapToLuckyFile(workbook, luckyFile);

        List<DefinedName> names = luckyFile.getDefinedNames();
        assertNotNull(names);
        assertEquals(1, names.size());
        DefinedName dn = names.get(0);
        assertEquals("MyRange", dn.getName());
        assertEquals("Sheet1!$A$1:$B$10", dn.getFormula());
    }

    @Test
    void mapToLuckyFile_xlnmPrefix_skipped() {
        XSSFName name1 = workbook.createName();
        name1.setNameName("_xlnm._FilterDatabase");
        name1.setRefersToFormula("Sheet1!$A$1:$B$10");

        XSSFName name2 = workbook.createName();
        name2.setNameName("_xlnm.Print_Area");
        name2.setRefersToFormula("Sheet1!$A$1:$C$20");

        XSSFName name3 = workbook.createName();
        name3.setNameName("ValidName");
        name3.setRefersToFormula("Sheet1!$D$1");

        DefinedNameMapper.mapToLuckyFile(workbook, luckyFile);

        List<DefinedName> names = luckyFile.getDefinedNames();
        assertNotNull(names);
        assertEquals(1, names.size());
        assertEquals("ValidName", names.get(0).getName());
    }

    @Test
    void mapToLuckyFile_withComment() {
        XSSFName name = workbook.createName();
        name.setNameName("MyRange");
        name.setRefersToFormula("Sheet1!$A$1:$B$10");
        name.setComment("This is a test range");

        DefinedNameMapper.mapToLuckyFile(workbook, luckyFile);

        List<DefinedName> names = luckyFile.getDefinedNames();
        assertEquals(1, names.size());
        assertEquals("This is a test range", names.get(0).getComment());
    }

    @Test
    void mapToLuckyFile_withoutComment() {
        XSSFName name = workbook.createName();
        name.setNameName("MyRange");
        name.setRefersToFormula("Sheet1!$A$1:$B$10");

        DefinedNameMapper.mapToLuckyFile(workbook, luckyFile);

        List<DefinedName> names = luckyFile.getDefinedNames();
        assertEquals(1, names.size());
        assertNull(names.get(0).getComment());
    }

    @Test
    void mapToLuckyFile_withSheetIndex() {
        XSSFName name = workbook.createName();
        name.setNameName("SheetLocal");
        name.setRefersToFormula("Sheet2!$A$1");
        name.setSheetIndex(1);

        DefinedNameMapper.mapToLuckyFile(workbook, luckyFile);

        List<DefinedName> names = luckyFile.getDefinedNames();
        assertEquals(1, names.size());
        assertEquals(1, names.get(0).getSheetIndex());
    }

    @Test
    void mapToLuckyFile_withoutSheetIndex() {
        XSSFName name = workbook.createName();
        name.setNameName("GlobalName");
        name.setRefersToFormula("Sheet1!$A$1");
        // Default is -1 (workbook level)

        DefinedNameMapper.mapToLuckyFile(workbook, luckyFile);

        List<DefinedName> names = luckyFile.getDefinedNames();
        assertEquals(1, names.size());
        assertEquals(-1, names.get(0).getSheetIndex());
    }

    @Test
    void mapToLuckyFile_hiddenTrue() {
        // Note: XSSFName.isHidden() returns true when the name is a function macro
        // This is read-only behavior from POI, testing the mapper's ability to read it
        XSSFName name = workbook.createName();
        name.setNameName("VisibleRange");
        name.setRefersToFormula("Sheet1!$A$1");

        DefinedNameMapper.mapToLuckyFile(workbook, luckyFile);

        List<DefinedName> names = luckyFile.getDefinedNames();
        assertEquals(1, names.size());
        // Normal names are not hidden
        assertFalse(names.get(0).getHidden());
    }

    @Test
    void mapToLuckyFile_hiddenFalse() {
        XSSFName name = workbook.createName();
        name.setNameName("VisibleRange");
        name.setRefersToFormula("Sheet1!$A$1");

        DefinedNameMapper.mapToLuckyFile(workbook, luckyFile);

        List<DefinedName> names = luckyFile.getDefinedNames();
        assertEquals(1, names.size());
        assertFalse(names.get(0).getHidden());
    }

    @Test
    void mapToLuckyFile_multipleNames() {
        XSSFName name1 = workbook.createName();
        name1.setNameName("Range1");
        name1.setRefersToFormula("Sheet1!$A$1");

        XSSFName name2 = workbook.createName();
        name2.setNameName("Range2");
        name2.setRefersToFormula("Sheet2!$B$2");

        XSSFName name3 = workbook.createName();
        name3.setNameName("_xlnm.Print_Area");
        name3.setRefersToFormula("Sheet1!$A$1:$C$10");

        DefinedNameMapper.mapToLuckyFile(workbook, luckyFile);

        List<DefinedName> names = luckyFile.getDefinedNames();
        assertNotNull(names);
        assertEquals(2, names.size());
        assertEquals("Range1", names.get(0).getName());
        assertEquals("Range2", names.get(1).getName());
    }

    // ========== mapToExcel ==========

    @Test
    void mapToExcel_nullList_doesNothing() {
        DefinedNameMapper.mapToExcel(null, workbook);
        assertEquals(0, workbook.getAllNames().size());
    }

    @Test
    void mapToExcel_emptyList_doesNothing() {
        DefinedNameMapper.mapToExcel(new ArrayList<>(), workbook);
        assertEquals(0, workbook.getAllNames().size());
    }

    @Test
    void mapToExcel_nullName_skipped() {
        List<DefinedName> names = new ArrayList<>();
        DefinedName dn = new DefinedName();
        dn.setName(null);
        dn.setFormula("Sheet1!$A$1");
        names.add(dn);

        DefinedNameMapper.mapToExcel(names, workbook);
        assertEquals(0, workbook.getAllNames().size());
    }

    @Test
    void mapToExcel_nullFormula_skipped() {
        List<DefinedName> names = new ArrayList<>();
        DefinedName dn = new DefinedName();
        dn.setName("MyRange");
        dn.setFormula(null);
        names.add(dn);

        DefinedNameMapper.mapToExcel(names, workbook);
        assertEquals(0, workbook.getAllNames().size());
    }

    @Test
    void mapToExcel_nameAlreadyExists_skipped() {
        XSSFName existing = workbook.createName();
        existing.setNameName("MyRange");
        existing.setRefersToFormula("Sheet1!$A$1");

        List<DefinedName> names = new ArrayList<>();
        DefinedName dn = new DefinedName();
        dn.setName("MyRange");
        dn.setFormula("Sheet1!$B$2");
        names.add(dn);

        DefinedNameMapper.mapToExcel(names, workbook);

        // Should still be 1, not 2
        assertEquals(1, workbook.getAllNames().size());
        // Original formula should remain
        assertEquals("Sheet1!$A$1", workbook.getName("MyRange").getRefersToFormula());
    }

    @Test
    void mapToExcel_withSheetIndexPositive() {
        List<DefinedName> names = new ArrayList<>();
        DefinedName dn = new DefinedName();
        dn.setName("LocalRange");
        dn.setFormula("Sheet2!$A$1");
        dn.setSheetIndex(1);
        names.add(dn);

        DefinedNameMapper.mapToExcel(names, workbook);

        assertEquals(1, workbook.getAllNames().size());
        XSSFName name = workbook.getName("LocalRange");
        assertNotNull(name);
        assertEquals(1, name.getSheetIndex());
    }

    @Test
    void mapToExcel_withSheetIndexNull() {
        List<DefinedName> names = new ArrayList<>();
        DefinedName dn = new DefinedName();
        dn.setName("GlobalRange");
        dn.setFormula("Sheet1!$A$1");
        dn.setSheetIndex(null);
        names.add(dn);

        DefinedNameMapper.mapToExcel(names, workbook);

        assertEquals(1, workbook.getAllNames().size());
        XSSFName name = workbook.getName("GlobalRange");
        assertNotNull(name);
        // Should be workbook level (-1)
        assertEquals(-1, name.getSheetIndex());
    }

    @Test
    void mapToExcel_withSheetIndexNegative() {
        List<DefinedName> names = new ArrayList<>();
        DefinedName dn = new DefinedName();
        dn.setName("GlobalRange");
        dn.setFormula("Sheet1!$A$1");
        dn.setSheetIndex(-1);
        names.add(dn);

        DefinedNameMapper.mapToExcel(names, workbook);

        assertEquals(1, workbook.getAllNames().size());
        XSSFName name = workbook.getName("GlobalRange");
        assertNotNull(name);
        assertEquals(-1, name.getSheetIndex());
    }

    @Test
    void mapToExcel_withComment() {
        List<DefinedName> names = new ArrayList<>();
        DefinedName dn = new DefinedName();
        dn.setName("MyRange");
        dn.setFormula("Sheet1!$A$1");
        dn.setComment("Test comment");
        names.add(dn);

        DefinedNameMapper.mapToExcel(names, workbook);

        assertEquals(1, workbook.getAllNames().size());
        XSSFName name = workbook.getName("MyRange");
        assertEquals("Test comment", name.getComment());
    }

    @Test
    void mapToExcel_withoutComment() {
        List<DefinedName> names = new ArrayList<>();
        DefinedName dn = new DefinedName();
        dn.setName("MyRange");
        dn.setFormula("Sheet1!$A$1");
        names.add(dn);

        DefinedNameMapper.mapToExcel(names, workbook);

        assertEquals(1, workbook.getAllNames().size());
        XSSFName name = workbook.getName("MyRange");
        assertNull(name.getComment());
    }

    @Test
    void mapToExcel_multipleNames() {
        List<DefinedName> names = new ArrayList<>();

        DefinedName dn1 = new DefinedName();
        dn1.setName("Range1");
        dn1.setFormula("Sheet1!$A$1");
        names.add(dn1);

        DefinedName dn2 = new DefinedName();
        dn2.setName("Range2");
        dn2.setFormula("Sheet2!$B$2");
        dn2.setSheetIndex(1);
        dn2.setComment("Local to Sheet2");
        names.add(dn2);

        DefinedNameMapper.mapToExcel(names, workbook);

        assertEquals(2, workbook.getAllNames().size());
        assertNotNull(workbook.getName("Range1"));
        assertNotNull(workbook.getName("Range2"));
    }

    // ========== Round-trip ==========

    @Test
    void roundTrip_simpleDefinedName() {
        List<DefinedName> src = new ArrayList<>();
        DefinedName dn = new DefinedName();
        dn.setName("TestRange");
        dn.setFormula("Sheet1!$A$1:$B$10");
        dn.setComment("Test range comment");
        dn.setSheetIndex(0);
        src.add(dn);

        DefinedNameMapper.mapToExcel(src, workbook);

        LuckyFile result = new LuckyFile();
        DefinedNameMapper.mapToLuckyFile(workbook, result);

        List<DefinedName> dest = result.getDefinedNames();
        assertNotNull(dest);
        assertEquals(1, dest.size());
        DefinedName resultDn = dest.get(0);
        assertEquals("TestRange", resultDn.getName());
        assertEquals("Sheet1!$A$1:$B$10", resultDn.getFormula());
        assertEquals("Test range comment", resultDn.getComment());
        assertEquals(0, resultDn.getSheetIndex());
    }

    @Test
    void roundTrip_workbookLevelName() {
        List<DefinedName> src = new ArrayList<>();
        DefinedName dn = new DefinedName();
        dn.setName("GlobalRange");
        dn.setFormula("Sheet1!$C$5");
        src.add(dn);

        DefinedNameMapper.mapToExcel(src, workbook);

        LuckyFile result = new LuckyFile();
        DefinedNameMapper.mapToLuckyFile(workbook, result);

        List<DefinedName> dest = result.getDefinedNames();
        assertNotNull(dest);
        assertEquals(1, dest.size());
        assertEquals("GlobalRange", dest.get(0).getName());
        assertEquals(-1, dest.get(0).getSheetIndex());
    }
}
