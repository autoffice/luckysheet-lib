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
package com.example.luckysheet.service;

import io.github.autoffice.luckysheet.LuckysheetConverter;
import io.github.autoffice.luckysheet.mapper.WorkbookMapperToExcel;
import io.github.autoffice.luckysheet.model.LuckyFile;
import io.github.autoffice.luckysheet.util.JSONUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Luckysheet 转换服务
 */
@Slf4j
@Service
public class LuckysheetService {

    /**
     * 上传 Excel 文件并转换为 Luckysheet JSON
     *
     * @param file 上传的 Excel 文件
     * @return Luckysheet JSON 字符串
     */
    public String uploadExcelAndConvert(MultipartFile file) throws IOException, InvalidFormatException {
        log.info("开始转换 Excel 文件: {}", file.getOriginalFilename());

        Path tempFile = Files.createTempFile("upload-", ".xlsx");
        try {
            // 先将 InputStream 完整写入临时文件再关闭，避免 Windows 文件锁
            try (InputStream is = file.getInputStream()) {
                Files.copy(is, tempFile, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
            }

            String jsonResult = LuckysheetConverter.excelToLuckySheetJson(tempFile.toString());

            log.info("Excel 转换成功，文件大小: {} bytes", file.getSize());
            return jsonResult;
        } finally {
            Files.deleteIfExists(tempFile);
        }
    }

    /**
     * 将 Luckysheet JSON 转换为 Excel 文件
     *
     * @param luckyFile Luckysheet 数据对象
     * @return Excel 文件字节数组
     */
    public byte[] convertToExcel(LuckyFile luckyFile) throws IOException {
        log.info("开始转换 Luckysheet JSON 为 Excel");

        XSSFWorkbook workbook = WorkbookMapperToExcel.mapToWorkbook(luckyFile);
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            workbook.write(out);
            workbook.close();
            byte[] excelBytes = out.toByteArray();
            log.info("Luckysheet 转换为 Excel 成功，文件大小: {} bytes", excelBytes.length);
            return excelBytes;
        }
    }

    /**
     * 获取示例 Luckysheet 数据 (从 classpath 中的 full.json 读取)
     *
     * @return 示例 JSON 字符串
     */
    public String getSampleData() throws IOException {
        ClassPathResource resource = new ClassPathResource("full.json");
        try (InputStream is = resource.getInputStream()) {
            byte[] bytes = new byte[is.available()];
            int offset = 0;
            while (offset < bytes.length) {
                int read = is.read(bytes, offset, bytes.length - offset);
                if (read < 0) {
                    break;
                }
                offset += read;
            }
            return new String(bytes, StandardCharsets.UTF_8);
        }
    }
}
