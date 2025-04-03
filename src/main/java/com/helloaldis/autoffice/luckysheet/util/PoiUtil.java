package com.helloaldis.autoffice.luckysheet.util;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class PoiUtil {
    public static int getFirstRowBase0(Sheet sheet) {
        int firstRowNum = sheet.getFirstRowNum();
        return Math.max(firstRowNum, 0);
    }

    public static short getMaxColNum(Sheet sheet) {
        short maxColNum = 0;
        for (int i = getFirstRowBase0(sheet); i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                maxColNum = (short) Math.max(maxColNum, row.getLastCellNum());
            }
        }

        return maxColNum;
    }
}
