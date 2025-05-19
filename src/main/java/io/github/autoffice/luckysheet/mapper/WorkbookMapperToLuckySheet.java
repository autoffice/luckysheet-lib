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

public class WorkbookMapperToLuckySheet {
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
                luckySheet.setOrder(i);
                SheetMapperToLuckySheet.mapToSheet(sheet, luckySheet);
                sheets.add(luckySheet);
            }
            luckyFile.setSheets(sheets);
        }

        return luckyFile;
    }
}
