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

/**
 * Apache POI 对象工厂.
 *
 * <p>集中创建 {@link XSSFWorkbook}、{@link XSSFCellStyle}、{@link Comment} 等 POI 对象,
 * 简化工作簿操作中的对象创建逻辑.</p>
 */
@Slf4j
public class PoiFactory {
    /**
     * 创建一个新的空 XSSFWorkbook 实例.
     *
     * @return 新创建的 XSSFWorkbook
     */
    public static XSSFWorkbook createWorkbook() {
        return new XSSFWorkbook();
    }

    /**
     * 静默关闭工作簿, 忽略 IO 异常.
     *
     * @param workbook 待关闭的工作簿, 可为 null
     */
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

    /**
     * 在指定工作簿中创建一个新的工作表.
     *
     * @param workbook 工作簿
     * @param name     工作表名称
     * @return 新创建的工作表
     */
    public static XSSFSheet createSheet(XSSFWorkbook workbook, String name) {
        return workbook.createSheet(name);
    }

    /**
     * 获取或创建指定行号的行 (若已存在则返回, 否则新建).
     *
     * @param sheet  工作表
     * @param rowNum 行号 (从 0 开始)
     * @return 对应的行实例
     */
    public static XSSFRow createOrGetRow(XSSFSheet sheet, int rowNum) {
        XSSFRow row = sheet.getRow(rowNum);
        if (row != null) {
            return row;
        }

        return sheet.createRow(rowNum);
    }

    /**
     * 获取或创建指定列号的单元格 (若已存在则返回, 否则新建).
     *
     * @param row    行
     * @param colNum 列号 (从 0 开始)
     * @return 对应的单元格实例
     */
    public static XSSFCell createOrGetCell(XSSFRow row, int colNum) {
        XSSFCell cell = row.getCell(colNum);
        if (cell != null) {
            return cell;
        }

        return row.createCell(colNum);
    }

    /**
     * 获取或创建指定行列的单元格 (若已存在则返回, 否则新建).
     *
     * @param sheet  工作表
     * @param rowNum 行号 (从 0 开始)
     * @param colNum 列号 (从 0 开始)
     * @return 对应的单元格实例
     */
    public static XSSFCell createOrGetCell(XSSFSheet sheet, int rowNum, int colNum) {
        XSSFRow row = createOrGetRow(sheet, rowNum);
        return createOrGetCell(row, colNum);
    }

    /**
     * 基于单元格所属工作簿创建新的单元格样式.
     *
     * @param cell 参考单元格 (用于获取工作簿)
     * @return 新的 XSSFCellStyle 实例
     */
    public static XSSFCellStyle createCellStyle(Cell cell) {
        return (XSSFCellStyle) cell.getSheet().getWorkbook().createCellStyle();
    }

    /**
     * 基于单元格所属工作簿创建数据格式对象.
     *
     * @param cell 参考单元格 (用于获取工作簿)
     * @return DataFormat 实例
     */
    public static DataFormat createDataFormat(Cell cell) {
        return cell.getSheet().getWorkbook().createDataFormat();
    }

    /**
     * 基于单元格所属工作簿创建新的字体对象.
     *
     * @param cell 参考单元格 (用于获取工作簿)
     * @return 新的 XSSFFont 实例
     */
    public static XSSFFont createFont(Cell cell) {
        return (XSSFFont) cell.getSheet().getWorkbook().createFont();
    }

    /**
     * 为指定单元格创建批注, 自动设置锚点位置.
     *
     * @param cell 目标单元格
     * @return 新创建的 Comment 实例
     */
    public static Comment createComment(Cell cell) {
        Drawing<?> drawingPatriarch = cell.getSheet().createDrawingPatriarch();
        ClientAnchor clientAnchor = cell.getSheet().getWorkbook().getCreationHelper().createClientAnchor();
        clientAnchor.setRow1(cell.getRowIndex());
        clientAnchor.setCol1(cell.getColumnIndex() + 2);
        clientAnchor.setRow2(cell.getRowIndex() + 6);
        clientAnchor.setCol2(cell.getColumnIndex() + 5);

        return drawingPatriarch.createCellComment(clientAnchor);
    }

    /**
     * 基于单元格所属工作簿创建空的富文本字符串对象.
     *
     * @param cell 参考单元格 (用于获取工作簿)
     * @return 新的 XSSFRichTextString 实例
     */
    public static XSSFRichTextString createRichTextString(XSSFCell cell) {
        return cell.getSheet().getWorkbook().getCreationHelper().createRichTextString("");
    }
}
