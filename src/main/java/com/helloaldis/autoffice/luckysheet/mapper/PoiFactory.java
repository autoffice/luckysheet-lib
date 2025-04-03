package com.helloaldis.autoffice.luckysheet.mapper;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

@Slf4j
public class PoiFactory {
    public static XSSFWorkbook createWorkbook() {
        return new XSSFWorkbook();
    }

    public static void closeWorkbookQuietly(XSSFWorkbook workbook) {
        if (workbook == null) {
            return;
        }

        try {
            workbook.close();
        } catch (IOException e) {
            log.error("close workbook error:", e);
        }
    }

    public static XSSFSheet createSheet(XSSFWorkbook workbook, String name) {
        return workbook.createSheet(name);
    }

    public static XSSFRow createOrGetRow(XSSFSheet sheet, int rowNum) {
        XSSFRow row = sheet.getRow(rowNum);
        if (row != null) {
            return row;
        }

        return sheet.createRow(rowNum);
    }

    public static XSSFCell createOrGetCell(XSSFRow row, int colNum) {
        XSSFCell cell = row.getCell(colNum);
        if (cell != null) {
            return cell;
        }

        return row.createCell(colNum);
    }

    public static XSSFCell createOrGetCell(XSSFSheet sheet, int rowNum, int colNum) {
        XSSFRow row = createOrGetRow(sheet, rowNum);
        return createOrGetCell(row, colNum);
    }

    public static XSSFCellStyle createCellStyle(Cell cell) {
        return (XSSFCellStyle) cell.getSheet().getWorkbook().createCellStyle();
    }

    public static XSSFFont createFont(Cell cell) {
        return (XSSFFont) cell.getSheet().getWorkbook().createFont();
    }

    public static Comment createComment(Cell cell) {
        Drawing<?> drawingPatriarch = cell.getSheet().createDrawingPatriarch();
        ClientAnchor clientAnchor = cell.getSheet().getWorkbook().getCreationHelper().createClientAnchor();
        clientAnchor.setRow1(cell.getRowIndex());
        clientAnchor.setCol1(cell.getColumnIndex() + 2);
        clientAnchor.setRow2(cell.getRowIndex() + 6);
        clientAnchor.setCol2(cell.getColumnIndex() + 5);

        return drawingPatriarch.createCellComment(clientAnchor);
    }

    public static XSSFRichTextString createRichTextString(XSSFCell cell) {
        return cell.getSheet().getWorkbook().getCreationHelper().createRichTextString("");
    }
}
