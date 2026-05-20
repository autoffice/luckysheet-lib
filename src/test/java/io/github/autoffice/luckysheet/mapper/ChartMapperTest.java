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

import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.model.sheet.Range;
import io.github.autoffice.luckysheet.model.sheet.SheetChart;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;

/**
 * 测试 {@link ChartMapper} 的双向映射逻辑, 覆盖各种图表类型和边界条件.
 */
class ChartMapperTest {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private LuckySheet luckySheet;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Test");
        luckySheet = new LuckySheet();
        luckySheet.setName("Test");
        luckySheet.setIndex("0");

        // 创建示例数据用于图表
        createSampleData(sheet);
    }

    @AfterEach
    void tearDown() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }

    private void createSampleData(XSSFSheet sheet) {
        // 表头
        sheet.createRow(0).createCell(0).setCellValue("Category");
        sheet.getRow(0).createCell(1).setCellValue("Value1");
        sheet.getRow(0).createCell(2).setCellValue("Value2");

        // 数据行
        sheet.createRow(1).createCell(0).setCellValue("A");
        sheet.getRow(1).createCell(1).setCellValue(10);
        sheet.getRow(1).createCell(2).setCellValue(20);

        sheet.createRow(2).createCell(0).setCellValue("B");
        sheet.getRow(2).createCell(1).setCellValue(15);
        sheet.getRow(2).createCell(2).setCellValue(25);

        sheet.createRow(3).createCell(0).setCellValue("C");
        sheet.getRow(3).createCell(1).setCellValue(12);
        sheet.getRow(3).createCell(2).setCellValue(22);
    }

    // ========== mapToLuckySheet tests ==========

    @Test
    void mapToLuckySheet_drawingNull_doesNothing() {
        ChartMapper.mapToLuckySheet(sheet, luckySheet);
        assertNull(luckySheet.getChart());
    }

    @Test
    void mapToLuckySheet_noCharts_doesNothing() {
        sheet.createDrawingPatriarch();
        ChartMapper.mapToLuckySheet(sheet, luckySheet);
        assertNull(luckySheet.getChart());
    }

    @Test
    void mapToLuckySheet_withBarChart() {
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 15);
        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("Bar Chart");

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                sheet, new CellRangeAddress(1, 3, 0, 0));
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                sheet, new CellRangeAddress(1, 3, 1, 1));

        XDDFBarChartData barData = (XDDFBarChartData) chart.createData(
                ChartTypes.BAR, bottomAxis, leftAxis);
        barData.addSeries(categories, values);
        chart.plot(barData);

        ChartMapper.mapToLuckySheet(sheet, luckySheet);

        assertNotNull(luckySheet.getChart());
        assertEquals(1, luckySheet.getChart().size());
        assertEquals("column", luckySheet.getChart().get(0).getChartOptions().getChartType());
        assertEquals("Bar Chart", luckySheet.getChart().get(0).getChartOptions().getTitle());
    }

    @Test
    void mapToLuckySheet_withLineChart() {
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 15);
        XSSFChart chart = drawing.createChart(anchor);

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);

        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                sheet, new CellRangeAddress(1, 3, 0, 0));
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                sheet, new CellRangeAddress(1, 3, 1, 1));

        XDDFLineChartData lineData = (XDDFLineChartData) chart.createData(
                ChartTypes.LINE, bottomAxis, leftAxis);
        lineData.addSeries(categories, values);
        chart.plot(lineData);

        ChartMapper.mapToLuckySheet(sheet, luckySheet);

        assertNotNull(luckySheet.getChart());
        assertEquals("line", luckySheet.getChart().get(0).getChartOptions().getChartType());
    }

    @Test
    void mapToLuckySheet_withPieChart() {
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 15);
        XSSFChart chart = drawing.createChart(anchor);

        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                sheet, new CellRangeAddress(1, 3, 0, 0));
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                sheet, new CellRangeAddress(1, 3, 1, 1));

        XDDFPieChartData pieData = (XDDFPieChartData) chart.createData(ChartTypes.PIE, null, null);
        pieData.addSeries(categories, values);
        chart.plot(pieData);

        ChartMapper.mapToLuckySheet(sheet, luckySheet);

        assertNotNull(luckySheet.getChart());
        assertEquals("pie", luckySheet.getChart().get(0).getChartOptions().getChartType());
    }

    @Test
    void mapToLuckySheet_titleTextNull() {
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 15);
        XSSFChart chart = drawing.createChart(anchor);

        XDDFPieChartData pieData = (XDDFPieChartData) chart.createData(ChartTypes.PIE, null, null);
        chart.plot(pieData);

        ChartMapper.mapToLuckySheet(sheet, luckySheet);

        assertNotNull(luckySheet.getChart());
        assertNull(luckySheet.getChart().get(0).getChartOptions().getTitle());
    }

    // ========== mapToExcel tests ==========

    @Test
    void mapToExcel_emptyChartList_doesNothing() {
        ChartMapper.mapToExcel(new ArrayList<>(), sheet);
        assertNull(sheet.getDrawingPatriarch());
    }

    @Test
    void mapToExcel_nullChartOptions() {
        SheetChart chart = new SheetChart();
        chart.setChartOptions(null);

        ChartMapper.mapToExcel(Collections.singletonList(chart), sheet);
        assertNotNull(sheet.getDrawingPatriarch());
    }

    @Test
    void mapToExcel_noRangeArray() {
        SheetChart chart = new SheetChart();
        SheetChart.ChartOptions opts = new SheetChart.ChartOptions();
        opts.setRangeArray(null);
        chart.setChartOptions(opts);

        ChartMapper.mapToExcel(Collections.singletonList(chart), sheet);
        assertNotNull(sheet.getDrawingPatriarch());
    }

    @Test
    void mapToExcel_invalidRange_firstRowGreaterOrEqualLastRow() {
        SheetChart chart = createBasicChart("column");
        chart.getChartOptions().setRangeArray(Collections.singletonList(buildRange(5, 5, 0, 2)));

        ChartMapper.mapToExcel(Collections.singletonList(chart), sheet);
        assertNotNull(sheet.getDrawingPatriarch());
    }

    @Test
    void mapToExcel_pieChart() {
        SheetChart chart = createBasicChart("pie");
        ChartMapper.mapToExcel(Collections.singletonList(chart), sheet);
        assertNotNull(sheet.getDrawingPatriarch());
        assertEquals(1, sheet.getDrawingPatriarch().getCharts().size());
    }

    @Test
    void mapToExcel_lineChart() {
        SheetChart chart = createBasicChart("line");
        ChartMapper.mapToExcel(Collections.singletonList(chart), sheet);
        assertNotNull(sheet.getDrawingPatriarch());
        assertEquals(1, sheet.getDrawingPatriarch().getCharts().size());
    }

    @Test
    void mapToExcel_columnChart() {
        SheetChart chart = createBasicChart("column");
        ChartMapper.mapToExcel(Collections.singletonList(chart), sheet);
        assertNotNull(sheet.getDrawingPatriarch());
        assertEquals(1, sheet.getDrawingPatriarch().getCharts().size());
    }

    @Test
    void mapToExcel_unknownChartType_defaultsToColumn() {
        SheetChart chart = createBasicChart("unknown");
        ChartMapper.mapToExcel(Collections.singletonList(chart), sheet);
        assertNotNull(sheet.getDrawingPatriarch());
        assertEquals(1, sheet.getDrawingPatriarch().getCharts().size());
    }

    @Test
    void mapToExcel_nullChartType_defaultsToColumn() {
        SheetChart chart = createBasicChart(null);
        ChartMapper.mapToExcel(Collections.singletonList(chart), sheet);
        assertNotNull(sheet.getDrawingPatriarch());
        assertEquals(1, sheet.getDrawingPatriarch().getCharts().size());
    }

    // ========== Helper methods ==========

    private SheetChart createBasicChart(String chartType) {
        SheetChart chart = new SheetChart();
        chart.setChart_id("chart1");

        SheetChart.ChartOptions opts = new SheetChart.ChartOptions();
        opts.setChartType(chartType);
        opts.setTitle("Test Chart");
        opts.setRangeArray(Collections.singletonList(buildRange(0, 3, 0, 2)));
        chart.setChartOptions(opts);

        return chart;
    }

    private Range buildRange(int r1, int r2, int c1, int c2) {
        Range r = new Range();
        r.setRow(Arrays.asList(r1, r2));
        r.setColumn(Arrays.asList(c1, c2));
        return r;
    }
}
