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
