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

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormat;
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

    public static DataFormat createDataFormat(Cell cell) {
        return cell.getSheet().getWorkbook().createDataFormat();
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
