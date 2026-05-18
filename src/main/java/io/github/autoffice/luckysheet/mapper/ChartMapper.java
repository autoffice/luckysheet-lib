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
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
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

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 图表 Luckysheet ↔ POI 双向映射器 (有损转换).
 *
 * <p>Excel 读出时只抽取类型、标题和数据范围等关键信息并放入 {@link SheetChart};
 * Luckysheet 写入时根据类型生成一个基础图表 (柱状、折线或饼图). 样式、主题等信息不保真.</p>
 */
public final class ChartMapper {

    private ChartMapper() {
    }

    public static void mapToLuckySheet(XSSFSheet sheet, LuckySheet luckySheet) {
        XSSFDrawing drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            return;
        }
        List<XSSFChart> charts = drawing.getCharts();
        if (charts == null || charts.isEmpty()) {
            return;
        }

        List<SheetChart> list = new ArrayList<>();
        for (XSSFChart xssfChart : charts) {
            SheetChart sheetChart = new SheetChart();
            sheetChart.setChart_id(xssfChart.getPackagePart().getPartName().getName());
            sheetChart.setSheetIndex(luckySheet.getIndex());

            SheetChart.ChartOptions opts = new SheetChart.ChartOptions();
            opts.setChart_id(sheetChart.getChart_id());
            opts.setChartType(detectChartType(xssfChart));
            if (xssfChart.getTitleText() != null) {
                opts.setTitle(xssfChart.getTitleText().getString());
            }
            opts.setRangeArray(extractRanges(xssfChart));
            sheetChart.setChartOptions(opts);
            list.add(sheetChart);
        }
        if (!list.isEmpty()) {
            luckySheet.setChart(list);
        }
    }

    public static void mapToExcel(List<SheetChart> charts, XSSFSheet sheet) {
        if (CollectionUtils.isEmpty(charts)) {
            return;
        }
        XSSFWorkbook workbook = sheet.getWorkbook();
        int sheetIndex = workbook.getSheetIndex(sheet);

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        int anchorOffset = 0;
        for (SheetChart chart : charts) {
            if (chart == null || chart.getChartOptions() == null) {
                continue;
            }
            SheetChart.ChartOptions opts = chart.getChartOptions();
            Range range = firstValidRange(opts.getRangeArray());
            if (range == null) {
                continue;
            }
            int firstRow = range.getRow().get(0);
            int lastRow = range.getRow().get(1);
            int firstCol = range.getColumn().get(0);
            int lastCol = range.getColumn().get(1);
            if (firstRow >= lastRow || firstCol >= lastCol) {
                continue;
            }

            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    lastCol + 1 + anchorOffset, firstRow,
                    lastCol + 8 + anchorOffset, firstRow + 15);
            anchorOffset += 8;

            XSSFChart xssfChart = drawing.createChart(anchor);
            if (opts.getTitle() != null) {
                xssfChart.setTitleText(opts.getTitle());
            }

            String type = opts.getChartType() == null ? "column" : opts.getChartType().toLowerCase();

            XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                    sheet, new CellRangeAddress(firstRow + 1, lastRow, firstCol, firstCol));
            List<XDDFNumericalDataSource<Double>> series = new ArrayList<>();
            for (int c = firstCol + 1; c <= lastCol; c++) {
                series.add(XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                        new CellRangeAddress(firstRow + 1, lastRow, c, c)));
            }

            if ("pie".equals(type)) {
                XDDFPieChartData pieData = (XDDFPieChartData) xssfChart.createData(ChartTypes.PIE, null, null);
                if (!series.isEmpty()) {
                    pieData.addSeries(categories, series.get(0));
                }
                xssfChart.plot(pieData);
            } else if ("line".equals(type)) {
                XDDFCategoryAxis bottomAxis = xssfChart.createCategoryAxis(AxisPosition.BOTTOM);
                XDDFValueAxis leftAxis = xssfChart.createValueAxis(AxisPosition.LEFT);
                leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
                XDDFLineChartData lineData = (XDDFLineChartData) xssfChart.createData(
                        ChartTypes.LINE, bottomAxis, leftAxis);
                for (XDDFNumericalDataSource<Double> s : series) {
                    XDDFChartData.Series seriesItem = lineData.addSeries(categories, s);
                    seriesItem.setTitle("Series", null);
                }
                xssfChart.plot(lineData);
            } else {
                XDDFCategoryAxis bottomAxis = xssfChart.createCategoryAxis(AxisPosition.BOTTOM);
                XDDFValueAxis leftAxis = xssfChart.createValueAxis(AxisPosition.LEFT);
                leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
                XDDFBarChartData barData = (XDDFBarChartData) xssfChart.createData(
                        ChartTypes.BAR, bottomAxis, leftAxis);
                for (XDDFNumericalDataSource<Double> s : series) {
                    XDDFChartData.Series seriesItem = barData.addSeries(categories, s);
                    seriesItem.setTitle("Series", null);
                }
                xssfChart.plot(barData);
            }
            XDDFChartLegend legend = xssfChart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);
            // 使 sheetIndex 在未来扩展中可用
            if (chart.getSheetIndex() == null) {
                chart.setSheetIndex(String.valueOf(sheetIndex));
            }
        }
    }

    private static String detectChartType(XSSFChart xssfChart) {
        List<XDDFChartData> dataList = xssfChart.getChartSeries();
        if (dataList == null || dataList.isEmpty()) {
            return "column";
        }
        XDDFChartData data = dataList.get(0);
        if (data instanceof XDDFPieChartData) {
            return "pie";
        }
        if (data instanceof XDDFLineChartData) {
            return "line";
        }
        if (data instanceof XDDFBarChartData) {
            return "column";
        }
        return "column";
    }

    private static List<Range> extractRanges(XSSFChart xssfChart) {
        List<Range> ranges = new ArrayList<>();
        List<XDDFChartData> dataList = xssfChart.getChartSeries();
        if (dataList == null) {
            return ranges;
        }
        for (XDDFChartData data : dataList) {
            for (int i = 0; i < data.getSeriesCount(); i++) {
                XDDFChartData.Series series = data.getSeries(i);
                String ref = null;
                try {
                    if (series.getValuesData() != null) {
                        ref = series.getValuesData().getFormula();
                    }
                } catch (Exception ignore) {
                    ref = null;
                }
                if (ref == null) {
                    continue;
                }
                // 截取 "SheetName!A1:B5" 的范围部分
                int bang = ref.indexOf('!');
                String addr = bang >= 0 ? ref.substring(bang + 1) : ref;
                addr = addr.replace("$", "");
                try {
                    CellRangeAddress cra = CellRangeAddress.valueOf(addr);
                    Range r = new Range();
                    r.setRow(Arrays.asList(cra.getFirstRow(), cra.getLastRow()));
                    r.setColumn(Arrays.asList(cra.getFirstColumn(), cra.getLastColumn()));
                    ranges.add(r);
                } catch (Exception ignore) {
                    // 非标准范围, 跳过
                    continue;
                }
            }
        }
        return ranges;
    }

    private static Range firstValidRange(List<Range> ranges) {
        if (ranges == null) {
            return null;
        }
        for (Range r : ranges) {
            if (r != null && r.getRow() != null && r.getColumn() != null
                    && r.getRow().size() >= 2 && r.getColumn().size() >= 2) {
                return r;
            }
        }
        return null;
    }
}
