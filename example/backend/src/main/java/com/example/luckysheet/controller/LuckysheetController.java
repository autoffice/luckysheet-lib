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
package com.example.luckysheet.controller;

import com.example.luckysheet.service.LuckysheetService;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.github.autoffice.luckysheet.model.LuckyFile;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

/**
 * Luckysheet 导入导出 REST API
 */
@Slf4j
@RestController
@RequestMapping("/api/excel")
@RequiredArgsConstructor
public class LuckysheetController {

    private final LuckysheetService luckysheetService;

    /**
     * 上传 Excel 文件并转换为 Luckysheet JSON
     *
     * @param file 上传的 Excel 文件
     * @return Luckysheet JSON 数据
     */
    @PostMapping(value = "/upload", produces = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<String> uploadExcel(@RequestParam("file") MultipartFile file) {
        try {
            if (file.isEmpty()) {
                return ResponseEntity.badRequest().body("{\"error\": \"文件不能为空\"}");
            }

            String originalFilename = file.getOriginalFilename();
            if (originalFilename == null || (!originalFilename.endsWith(".xlsx") && !originalFilename.endsWith(".xls"))) {
                return ResponseEntity.badRequest().body("{\"error\": \"只支持 Excel 文件 (.xlsx, .xls)\"}");
            }

            String jsonResult = luckysheetService.uploadExcelAndConvert(file);
            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_JSON)
                    .body(jsonResult);

        } catch (IOException | InvalidFormatException e) {
            log.error("Excel 文件转换失败", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("{\"error\": \"文件转换失败: " + e.getMessage() + "\"}");
        }
    }

    /**
     * 接收 Luckysheet JSON 并导出为 Excel 文件
     *
     * @param jsonBody Luckysheet JSON 字符串
     * @return Excel 文件
     */
    @PostMapping(value = "/export", consumes = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<byte[]> exportExcel(@RequestBody String jsonBody) {
        try {
            ObjectMapper mapper = new ObjectMapper();
            mapper.configure(com.fasterxml.jackson.databind.DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false);
            LuckyFile luckyFile = mapper.readValue(jsonBody, LuckyFile.class);

            byte[] excelBytes = luckysheetService.convertToExcel(luckyFile);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData("attachment",
                    URLEncoder.encode("导出文件.xlsx", "UTF-8"));
            headers.setContentLength(excelBytes.length);

            return ResponseEntity.ok()
                    .headers(headers)
                    .body(excelBytes);

        } catch (Exception e) {
            log.error("导出 Excel 失败", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body(("{\"error\": \"导出失败: " + e.getMessage() + "\"}").getBytes());
        }
    }

    /**
     * 获取示例 Luckysheet 数据
     *
     * @return 示例 JSON 数据
     */
    @GetMapping(value = "/sample", produces = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<String> getSampleData() {
        try {
            String sampleData = luckysheetService.getSampleData();
            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_JSON)
                    .body(sampleData);
        } catch (IOException e) {
            log.error("加载示例数据失败", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("{\"error\": \"加载示例数据失败: " + e.getMessage() + "\"}");
        }
    }

    /**
     * 健康检查接口
     */
    @GetMapping("/health")
    public ResponseEntity<String> health() {
        return ResponseEntity.ok("{\"status\": \"UP\"}");
    }
}
