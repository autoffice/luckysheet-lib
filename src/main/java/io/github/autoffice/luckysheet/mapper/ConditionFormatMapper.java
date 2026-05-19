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

import io.github.autoffice.luckysheet.model.sheet.ConditionFormat;
import io.github.autoffice.luckysheet.model.sheet.ConditionFormatType;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.model.sheet.Range;
import io.github.autoffice.luckysheet.util.NumberUtil;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFColorScaleFormatting;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFDataBarFormatting;
import org.apache.poi.xssf.usermodel.XSSFIconMultiStateFormatting;
import org.apache.poi.xssf.usermodel.XSSFPatternFormatting;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 条件格式 Luckysheet ↔ POI 双向映射器 (有损: 复杂规则可能无法完整还原).
 */
public final class ConditionFormatMapper {

    private ConditionFormatMapper() {
    }

    /**
     * 从 Excel 工作表中提取条件格式并转换为 Luckysheet 格式.
     *
     * @param sheet      源 POI 工作表
     * @param luckySheet 目标 Luckysheet 工作表
     */
    public static void mapToLuckySheet(XSSFSheet sheet, LuckySheet luckySheet) {
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        int count = scf.getNumConditionalFormattings();
        if (count == 0) {
            return;
        }
        List<ConditionFormat> saves = luckySheet.getLuckysheet_conditionformat_save();
        if (saves == null) {
            saves = new ArrayList<>();
            luckySheet.setLuckysheet_conditionformat_save(saves);
        }

        for (int i = 0; i < count; i++) {
            ConditionalFormatting cf = scf.getConditionalFormattingAt(i);
            CellRangeAddress[] ranges = cf.getFormattingRanges();
            if (ranges == null || ranges.length == 0) {
                continue;
            }
            List<Range> modelRanges = new ArrayList<>();
            for (CellRangeAddress address : ranges) {
                Range r = new Range();
                r.setRow(Arrays.asList(address.getFirstRow(), address.getLastRow()));
                r.setColumn(Arrays.asList(address.getFirstColumn(), address.getLastColumn()));
                modelRanges.add(r);
            }
            for (int j = 0; j < cf.getNumberOfRules(); j++) {
                ConditionalFormattingRule rule = cf.getRule(j);
                ConditionFormat saved = buildFromPoiRule(rule, modelRanges);
                if (saved != null) {
                    saves.add(saved);
                }
            }
        }
    }

    /**
     * 将 Luckysheet 条件格式列表写入 Excel 工作表.
     *
     * @param saves 条件格式列表
     * @param sheet 目标 POI 工作表
     */
    public static void mapToExcel(List<ConditionFormat> saves, XSSFSheet sheet) {
        if (CollectionUtils.isEmpty(saves)) {
            return;
        }
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        for (ConditionFormat cf : saves) {
            if (cf == null || CollectionUtils.isEmpty(cf.getCellrange()) || cf.getType() == null) {
                continue;
            }
            CellRangeAddress[] addresses = buildRangeAddresses(cf.getCellrange());
            if (addresses.length == 0) {
                continue;
            }
            XSSFConditionalFormattingRule rule = buildPoiRule(scf, cf);
            if (rule == null) {
                continue;
            }
            scf.addConditionalFormatting(addresses, rule);
        }
    }

    private static ConditionFormat buildFromPoiRule(ConditionalFormattingRule rule, List<Range> ranges) {
        ConditionFormat model = new ConditionFormat();
        model.setCellrange(ranges);

        ConditionType ct = rule.getConditionType();
        if (ct == ConditionType.DATA_BAR) {
            model.setType(ConditionFormatType.DATA_BAR);
            Map<String, Object> fmt = new HashMap<>();
            if (rule instanceof XSSFConditionalFormattingRule) {
                XSSFDataBarFormatting db = ((XSSFConditionalFormattingRule) rule).getDataBarFormatting();
                if (db != null && db.getColor() != null) {
                    fmt.put("color", NumberUtil.rgbToColorString(db.getColor().getRGB()));
                }
            }
            model.setFormat(fmt);
            model.setConditionName("dataBar");
            return model;
        }
        if (ct == ConditionType.COLOR_SCALE) {
            model.setType(ConditionFormatType.COLOR_GRADATION);
            Map<String, Object> fmt = new HashMap<>();
            if (rule instanceof XSSFConditionalFormattingRule) {
                XSSFColorScaleFormatting cs = ((XSSFConditionalFormattingRule) rule).getColorScaleFormatting();
                if (cs != null && cs.getColors() != null) {
                    XSSFColor[] colors = cs.getColors();
                    if (colors.length > 0 && colors[0] != null) {
                        fmt.put("leastcolor", NumberUtil.rgbToColorString(colors[0].getRGB()));
                    }
                    if (colors.length == 3 && colors[1] != null) {
                        fmt.put("middlecolor", NumberUtil.rgbToColorString(colors[1].getRGB()));
                    }
                    if (colors.length >= 2 && colors[colors.length - 1] != null) {
                        fmt.put("maxcolor", NumberUtil.rgbToColorString(colors[colors.length - 1].getRGB()));
                    }
                }
            }
            model.setFormat(fmt);
            model.setConditionName("colorGradation");
            return model;
        }
        if (ct == ConditionType.ICON_SET) {
            model.setType(ConditionFormatType.ICONS);
            model.setFormat(Collections.singletonMap("leasticons", "3Arrows"));
            model.setConditionName("icons");
            return model;
        }

        // default: cell value comparison
        model.setType(ConditionFormatType.DEFAULT);
        Map<String, Object> fmt = new HashMap<>();
        if (rule instanceof XSSFConditionalFormattingRule) {
            XSSFPatternFormatting pf = ((XSSFConditionalFormattingRule) rule).getPatternFormatting();
            if (pf != null && pf.getFillBackgroundColorColor() instanceof XSSFColor) {
                XSSFColor fillColor = (XSSFColor) pf.getFillBackgroundColorColor();
                byte[] rgb = fillColor.getRGB();
                if (rgb != null) {
                    fmt.put("cellColor", NumberUtil.rgbToColorString(rgb));
                }
            }
        }
        fmt.putIfAbsent("textColor", "#000000");
        fmt.putIfAbsent("cellColor", "#ffff00");
        model.setFormat(fmt);
        model.setConditionName(mapOperatorToConditionName(rule.getComparisonOperation()));
        String f1 = rule.getFormula1();
        String f2 = rule.getFormula2();
        List<Object> cond = new ArrayList<>();
        if (f1 != null) {
            cond.add(f1);
        }
        if (f2 != null) {
            cond.add(f2);
        }
        if (!cond.isEmpty()) {
            model.setConditionValue(cond);
        }
        return model;
    }

    private static XSSFConditionalFormattingRule buildPoiRule(SheetConditionalFormatting scf, ConditionFormat cf) {
        switch (cf.getType()) {
            case DATA_BAR: {
                XSSFColor color = new XSSFColor(NumberUtil.colorStringToRgb(resolveFormatColor(cf.getFormat(), "color", "#638EC6")), null);
                return (XSSFConditionalFormattingRule) scf.createConditionalFormattingRule(color);
            }
            case COLOR_GRADATION: {
                XSSFColor[] colors = new XSSFColor[3];
                colors[0] = new XSSFColor(NumberUtil.colorStringToRgb(resolveFormatColor(cf.getFormat(), "leastcolor", "#FFEF9C")), null);
                colors[1] = new XSSFColor(NumberUtil.colorStringToRgb(resolveFormatColor(cf.getFormat(), "middlecolor", "#FCFCFF")), null);
                colors[2] = new XSSFColor(NumberUtil.colorStringToRgb(resolveFormatColor(cf.getFormat(), "maxcolor", "#63BE7B")), null);
                XSSFConditionalFormattingRule rule = (XSSFConditionalFormattingRule) scf.createConditionalFormattingColorScaleRule();
                XSSFColorScaleFormatting csf = rule.getColorScaleFormatting();
                csf.setNumControlPoints(3);
                csf.setColors(colors);
                ConditionalFormattingThreshold[] thresholds = csf.getThresholds();
                thresholds[0].setRangeType(ConditionalFormattingThreshold.RangeType.MIN);
                thresholds[1].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENTILE);
                thresholds[1].setValue(50d);
                thresholds[2].setRangeType(ConditionalFormattingThreshold.RangeType.MAX);
                csf.setThresholds(thresholds);
                return rule;
            }
            case ICONS: {
                XSSFConditionalFormattingRule rule = (XSSFConditionalFormattingRule) scf.createConditionalFormattingRule(
                        IconMultiStateFormatting.IconSet.GYR_3_ARROW);
                XSSFIconMultiStateFormatting icons = rule.getMultiStateFormatting();
                if (icons != null) {
                    ConditionalFormattingThreshold[] thresholds = icons.getThresholds();
                    if (thresholds != null && thresholds.length >= 3) {
                        thresholds[0].setRangeType(ConditionalFormattingThreshold.RangeType.MIN);
                        thresholds[1].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
                        thresholds[1].setValue(33d);
                        thresholds[2].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
                        thresholds[2].setValue(67d);
                        icons.setThresholds(thresholds);
                    }
                }
                return rule;
            }
            case DEFAULT:
            default: {
                byte op = mapConditionNameToOperator(cf.getConditionName());
                String[] formulas = resolveFormulas(cf.getConditionValue());
                XSSFConditionalFormattingRule rule;
                if (formulas.length >= 2) {
                    rule = (XSSFConditionalFormattingRule) scf.createConditionalFormattingRule(op, formulas[0], formulas[1]);
                } else if (formulas.length == 1) {
                    rule = (XSSFConditionalFormattingRule) scf.createConditionalFormattingRule(op, formulas[0]);
                } else {
                    rule = (XSSFConditionalFormattingRule) scf.createConditionalFormattingRule(op, "0");
                }
                String color = resolveFormatColor(cf.getFormat(), "cellColor", "#ffff00");
                XSSFPatternFormatting pf = rule.createPatternFormatting();
                pf.setFillBackgroundColor(new XSSFColor(NumberUtil.colorStringToRgb(color), null));
                pf.setFillPattern(XSSFPatternFormatting.SOLID_FOREGROUND);
                return rule;
            }
        }
    }

    private static byte mapConditionNameToOperator(String name) {
        if (name == null) {
            return ComparisonOperator.BETWEEN;
        }
        switch (name) {
            case "betweenness":
                return ComparisonOperator.BETWEEN;
            case "notBetweenness":
                return ComparisonOperator.NOT_BETWEEN;
            case "greatThan":
            case "greaterThan":
                return ComparisonOperator.GT;
            case "greatEqual":
                return ComparisonOperator.GE;
            case "lessThan":
                return ComparisonOperator.LT;
            case "lessEqual":
                return ComparisonOperator.LE;
            case "equal":
                return ComparisonOperator.EQUAL;
            case "notEqual":
                return ComparisonOperator.NOT_EQUAL;
            default:
                return ComparisonOperator.BETWEEN;
        }
    }

    private static String mapOperatorToConditionName(byte op) {
        switch (op) {
            case ComparisonOperator.BETWEEN:
                return "betweenness";
            case ComparisonOperator.NOT_BETWEEN:
                return "notBetweenness";
            case ComparisonOperator.GT:
                return "greatThan";
            case ComparisonOperator.GE:
                return "greatEqual";
            case ComparisonOperator.LT:
                return "lessThan";
            case ComparisonOperator.LE:
                return "lessEqual";
            case ComparisonOperator.EQUAL:
                return "equal";
            case ComparisonOperator.NOT_EQUAL:
                return "notEqual";
            default:
                return null;
        }
    }

    private static String[] resolveFormulas(List<Object> values) {
        if (CollectionUtils.isEmpty(values)) {
            return new String[0];
        }
        String[] arr = new String[values.size()];
        for (int i = 0; i < values.size(); i++) {
            Object o = values.get(i);
            arr[i] = o == null ? "" : String.valueOf(o);
        }
        return arr;
    }

    private static String resolveFormatColor(Object format, String key, String defaultValue) {
        if (format instanceof Map) {
            Object v = ((Map<?, ?>) format).get(key);
            if (v != null) {
                return String.valueOf(v);
            }
        }
        return defaultValue;
    }

    private static CellRangeAddress[] buildRangeAddresses(List<Range> ranges) {
        List<CellRangeAddress> list = new ArrayList<>();
        for (Range r : ranges) {
            if (r == null || r.getRow() == null || r.getColumn() == null
                    || r.getRow().size() < 2 || r.getColumn().size() < 2) {
                continue;
            }
            list.add(new CellRangeAddress(r.getRow().get(0), r.getRow().get(1),
                    r.getColumn().get(0), r.getColumn().get(1)));
        }
        return list.toArray(new CellRangeAddress[0]);
    }
}
