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
import io.github.autoffice.luckysheet.model.LuckyFileInfo;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.util.DateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

/**
 * 工作簿 POI → Luckysheet 映射器.
 *
 * <p>负责将 xlsx 文件读取为 POI {@link XSSFWorkbook} 后转换为 {@link LuckyFile},
 * 包括所有工作表以及工作簿级命名范围.</p>
 */
public class WorkbookMapperToLuckySheet {
    /**
     * 读取 xlsx 文件并转换为 Luckysheet 文件对象.
     *
     * @param src xlsx 文件路径
     * @return 解析后的 Luckysheet 文件对象
     * @throws IOException            IO 异常
     * @throws InvalidFormatException xlsx 文件格式不合法
     */
    public static LuckyFile mapToLuckyFile(String src) throws IOException, InvalidFormatException {
        LuckyFile luckyFile = LuckySheetFactory.createLuckyFile();
        LuckyFileInfo luckyFileInfo = LuckySheetFactory.createLuckyFileInfo();
        luckyFile.setInfo(luckyFileInfo);

        File file = new File(src);
        luckyFileInfo.setName(file.getName().replaceAll(".xlsx", ""));
        luckyFileInfo.setCreatedTime(DateUtil.toJsonTimeString(System.currentTimeMillis()));
        luckyFileInfo.setModifiedTime(DateUtil.toJsonTimeString(System.currentTimeMillis()));

        try (XSSFWorkbook workbook = new XSSFWorkbook(file)) {
            ArrayList<LuckySheet> sheets = new ArrayList<>(workbook.getNumberOfSheets());
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                LuckySheet luckySheet = LuckySheetFactory.createLuckySheet(sheet.getSheetName());
                luckySheet.setIndex(String.valueOf(i));
                if (sheet.getTabColor() != null) {
                    luckySheet.setColor("#" + sheet.getTabColor().getARGBHex().substring(2));
                }
                luckySheet.setOrder(i);
                SheetMapperToLuckySheet.mapToSheet(sheet, luckySheet);
                sheets.add(luckySheet);
            }
            luckyFile.setSheets(sheets);
            DefinedNameMapper.mapToLuckyFile(workbook, luckyFile);
        }

        return luckyFile;
    }
}
