package com.helloaldis.autoffice.luckysheet;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.net.URL;

import static org.junit.jupiter.api.Assertions.assertNotNull;

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
        LuckysheetConverter.luckysheetToExcel(OUTPUT + "full2.json", OUTPUT + "full2.xlsx");
    }
}
