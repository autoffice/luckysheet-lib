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
