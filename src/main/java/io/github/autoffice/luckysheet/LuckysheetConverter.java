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

import io.github.autoffice.luckysheet.mapper.PoiFactory;
import io.github.autoffice.luckysheet.mapper.WorkbookMapperToExcel;
import io.github.autoffice.luckysheet.mapper.WorkbookMapperToLuckySheet;
import io.github.autoffice.luckysheet.model.LuckyFile;
import io.github.autoffice.luckysheet.util.JSONUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;

@Slf4j
public class LuckysheetConverter {

    /**
     * luckysheet json 文件转为xlsx文件
     *
     * @param srcLuckySheet luckysheet json 文件路径
     * @param destXlsx      xlsx文件路径
     * @throws IOException 转换异常
     */
    public static void luckysheetToExcel(String srcLuckySheet, String destXlsx) throws IOException {
            LuckyFile luckyFile = readAsLuckyFile(srcLuckySheet);
        XSSFWorkbook workbook = null;
        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(destXlsx);
            workbook = WorkbookMapperToExcel.mapToWorkbook(luckyFile);
            workbook.write(outputStream);
        } finally {
            PoiFactory.closeWorkbookQuietly(workbook);
            IOUtils.closeQuietly(outputStream);
        }
    }

    /**
     * luckysheet json 文件转为xlsx文件流
     *
     * @param srcLuckySheet    luckysheet json 文件路径
     * @param xlsxOutputStream xlsx文件流
     * @throws IOException 转换异常
     */
    @SuppressWarnings("unused")
    public static void luckysheetToExcel(String srcLuckySheet, OutputStream xlsxOutputStream) throws IOException {
        LuckyFile luckyFile = readAsLuckyFile(srcLuckySheet);
        XSSFWorkbook workbook = null;
        try {
            workbook = WorkbookMapperToExcel.mapToWorkbook(luckyFile);
            workbook.write(xlsxOutputStream);
        } finally {
            PoiFactory.closeWorkbookQuietly(workbook);
            IOUtils.closeQuietly(xlsxOutputStream);
        }
    }

    /**
     * xlsx文件转为luckysheet json文件
     *
     * @param srcXlsx        xlsx文件路径
     * @param destLuckySheet luckysheet json文件路径
     * @throws IOException            转换异常
     * @throws InvalidFormatException 转换异常
     */
    public static void excelToLuckySheetFile(String srcXlsx, String destLuckySheet) throws IOException, InvalidFormatException {
        FileUtils.writeStringToFile(new File(destLuckySheet), excelToLuckySheetJson(srcXlsx), StandardCharsets.UTF_8);
    }

    /**
     * xlsx 文件转为luckysheet json string
     *
     * @param srcXlsx xlsx文件路径
     * @return luckysheet json string
     * @throws IOException            转换异常
     * @throws InvalidFormatException 转换异常
     */
    public static String excelToLuckySheetJson(String srcXlsx) throws IOException, InvalidFormatException {
        return JSONUtil.asJSONString(excelToLuckySheet(srcXlsx));
    }

    /**
     * xlsx 文件转为luckysheet 对象
     *
     * @param srcXlsx xlsx文件路径
     * @return luckysheet 对象
     * @throws IOException            转换异常
     * @throws InvalidFormatException 转换异常
     */
    public static LuckyFile excelToLuckySheet(String srcXlsx) throws IOException, InvalidFormatException {
        return WorkbookMapperToLuckySheet.mapToLuckyFile(srcXlsx);
    }

    /**
     * 读取json文件为luckysheet对象
     *
     * @param srcJsonFile json文件路径
     * @return luckysheet 对象
     */
    public static LuckyFile readAsLuckyFile(String srcJsonFile) {
        return JSONUtil.asEntity(new File(srcJsonFile), LuckyFile.class);
    }
}
