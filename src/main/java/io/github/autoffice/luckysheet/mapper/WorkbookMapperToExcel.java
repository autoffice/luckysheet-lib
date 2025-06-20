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

import io.github.autoffice.luckysheet.model.LuckyFile;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkbookMapperToExcel {
    /**
     * 将luckysheet对象转换为POI中的workbook
     *
     * @param luckyFile luckysheet对象
     * @return POI workbook对象
     */
    public static XSSFWorkbook mapToWorkbook(LuckyFile luckyFile) {
        XSSFWorkbook workbook = PoiFactory.createWorkbook();

        // 一个workbook包含多个sheet，遍历并转换
        for (LuckySheet luckySheet : luckyFile.getSheets()) {
            XSSFSheet poiSheet = PoiFactory.createSheet(workbook, luckySheet.getName());
            SheetMapperToExcel.mapToSheet(luckySheet, poiSheet);
        }

        return workbook;
    }
}
