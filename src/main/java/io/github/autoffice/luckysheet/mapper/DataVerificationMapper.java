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

import io.github.autoffice.luckysheet.model.sheet.DataVerification;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 数据验证 Luckysheet ↔ POI 双向映射器.
 */
@Slf4j
public final class DataVerificationMapper {

    private DataVerificationMapper() {
    }

    /**
     * 从 Excel 工作表中提取数据验证规则并转换为 Luckysheet 格式.
     *
     * @param sheet      源 POI 工作表
     * @param luckySheet 目标 Luckysheet 工作表
     */
    public static void mapToLuckySheet(XSSFSheet sheet, LuckySheet luckySheet) {
        List<XSSFDataValidation> validations = sheet.getDataValidations();
        if (validations == null || validations.isEmpty()) {
            return;
        }

        Map<String, DataVerification> map = luckySheet.getDataVerification();
        if (map == null) {
            map = new LinkedHashMap<>();
            luckySheet.setDataVerification(map);
        }

        for (XSSFDataValidation dv : validations) {
            DataVerification model = buildFromPoi(dv);
            if (model == null) {
                continue;
            }
            CellRangeAddressList regions = dv.getRegions();
            if (regions == null) {
                continue;
            }
            for (CellRangeAddress range : regions.getCellRangeAddresses()) {
                for (int r = range.getFirstRow(); r <= range.getLastRow(); r++) {
                    for (int c = range.getFirstColumn(); c <= range.getLastColumn(); c++) {
                        map.put(r + "_" + c, copy(model));
                    }
                }
            }
        }
    }

    /**
     * 将 Luckysheet 数据验证规则写入 Excel 工作表.
     *
     * @param verifications 数据验证映射 (key 为 "行_列" 格式)
     * @param sheet         目标 POI 工作表
     */
    public static void mapToExcel(Map<String, DataVerification> verifications, XSSFSheet sheet) {
        if (MapUtils.isEmpty(verifications)) {
            return;
        }

        XSSFDataValidationHelper helper = new XSSFDataValidationHelper(sheet);
        // 合并相同规则到同一 region, 提升 xlsx 体积效率
        Map<DataVerification, CellRangeAddressList> grouped = new HashMap<>();
        for (Map.Entry<String, DataVerification> entry : verifications.entrySet()) {
            int[] rc = parseRowColKey(entry.getKey());
            if (rc == null) {
                log.warn("Invalid data verification key: {}", entry.getKey());
                continue;
            }
            if (entry.getValue() == null) {
                continue;
            }
            CellRangeAddress cra = new CellRangeAddress(rc[0], rc[0], rc[1], rc[1]);
            CellRangeAddressList list = grouped.computeIfAbsent(entry.getValue(), k -> new CellRangeAddressList());
            list.addCellRangeAddress(cra);
        }

        for (Map.Entry<DataVerification, CellRangeAddressList> entry : grouped.entrySet()) {
            DataValidationConstraint constraint = buildConstraint(helper, entry.getKey());
            if (constraint == null) {
                log.warn("Unsupported data verification type: {}, type2: {}", entry.getKey().getType(), entry.getKey().getType2());
                continue;
            }
            DataValidation validation = helper.createValidation(constraint, entry.getValue());
            DataVerification v = entry.getKey();
            if (Boolean.TRUE.equals(v.getProhibitInput())) {
                validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
                validation.setShowErrorBox(true);
                validation.createErrorBox("", "");
            }
            if (Boolean.TRUE.equals(v.getHintShow())) {
                validation.setShowPromptBox(true);
                validation.createPromptBox("", StringUtils.defaultString(v.getHintText()));
            }
            validation.setSuppressDropDownArrow(isListType(v.getType()));
            sheet.addValidationData(validation);
        }
    }

    private static DataVerification buildFromPoi(XSSFDataValidation dv) {
        DataValidationConstraint constraint = dv.getValidationConstraint();
        if (constraint == null) {
            return null;
        }
        DataVerification model = new DataVerification();
        int validationType = constraint.getValidationType();
        int op = constraint.getOperator();
        String formula1 = constraint.getFormula1();
        String formula2 = constraint.getFormula2();
        model.setValue1(formula1);
        model.setValue2(formula2);
        switch (validationType) {
            case DataValidationConstraint.ValidationType.LIST:
                model.setType("dropdown");
                if (formula1 != null) {
                    String[] explicit = constraint.getExplicitListValues();
                    if (explicit != null && explicit.length > 0) {
                        model.setValue1(String.join(",", explicit));
                    }
                }
                break;
            case DataValidationConstraint.ValidationType.INTEGER:
            case DataValidationConstraint.ValidationType.DECIMAL:
                model.setType("number");
                model.setType2(mapOperatorToLuckyType2(op, false));
                break;
            case DataValidationConstraint.ValidationType.TEXT_LENGTH:
                model.setType("text_length");
                model.setType2(mapOperatorToLuckyType2(op, false));
                break;
            case DataValidationConstraint.ValidationType.DATE:
                model.setType("date");
                model.setType2(mapOperatorToLuckyType2(op, true));
                break;
            case DataValidationConstraint.ValidationType.FORMULA:
                model.setType("text_content");
                model.setType2("include");
                break;
            default:
                model.setType("text_content");
                break;
        }
        model.setProhibitInput(dv.getErrorStyle() == DataValidation.ErrorStyle.STOP);
        model.setHintShow(dv.getShowPromptBox());
        model.setHintText(dv.getPromptBoxText());
        model.setChecked(false);
        model.setRemote(false);
        return model;
    }

    private static DataValidationConstraint buildConstraint(DataValidationHelper helper, DataVerification v) {
        String type = StringUtils.lowerCase(StringUtils.defaultString(v.getType()));
        String value1 = StringUtils.defaultString(v.getValue1());
        String value2 = StringUtils.defaultString(v.getValue2());
        switch (type) {
            case "dropdown":
                return helper.createExplicitListConstraint(splitList(value1));
            case "checkbox":
                return helper.createExplicitListConstraint(new String[]{"TRUE", "FALSE"});
            case "number": {
                int op = mapType2ToOperator(v.getType2());
                return helper.createDecimalConstraint(op, value1, value2);
            }
            case "text_length": {
                int op = mapType2ToOperator(v.getType2());
                return helper.createTextLengthConstraint(op, value1, value2);
            }
            case "date": {
                int op = mapType2ToOperator(v.getType2());
                return helper.createDateConstraint(op, value1, value2, "yyyy-MM-dd");
            }
            case "text_content":
                return helper.createCustomConstraint(buildTextContentFormula(v));
            case "validity":
                return helper.createCustomConstraint(buildValidityFormula(v));
            default:
                return null;
        }
    }

    private static String buildTextContentFormula(DataVerification v) {
        String t2 = StringUtils.defaultString(v.getType2(), "include");
        String val = StringUtils.defaultString(v.getValue1());
        if ("include".equals(t2)) {
            return "ISNUMBER(SEARCH(\"" + escape(val) + "\",INDIRECT(ADDRESS(ROW(),COLUMN()))))";
        } else if ("exclude".equals(t2)) {
            return "ISERROR(SEARCH(\"" + escape(val) + "\",INDIRECT(ADDRESS(ROW(),COLUMN()))))";
        } else if ("equal".equals(t2)) {
            return "INDIRECT(ADDRESS(ROW(),COLUMN()))=\"" + escape(val) + "\"";
        }
        return "TRUE";
    }

    private static String buildValidityFormula(DataVerification v) {
        String t2 = StringUtils.defaultString(v.getType2());
        if ("phone".equalsIgnoreCase(t2)) {
            return "AND(LEN(INDIRECT(ADDRESS(ROW(),COLUMN())))=11,ISNUMBER(VALUE(INDIRECT(ADDRESS(ROW(),COLUMN())))))";
        }
        if ("card".equalsIgnoreCase(t2)) {
            return "OR(LEN(INDIRECT(ADDRESS(ROW(),COLUMN())))=15,LEN(INDIRECT(ADDRESS(ROW(),COLUMN())))=18)";
        }
        return "TRUE";
    }

    private static String escape(String s) {
        return s == null ? "" : s.replace("\"", "\"\"");
    }

    private static String[] splitList(String raw) {
        if (StringUtils.isBlank(raw)) {
            return new String[0];
        }
        String[] parts = raw.split(",");
        for (int i = 0; i < parts.length; i++) {
            parts[i] = parts[i].trim();
        }
        return parts;
    }

    private static int mapType2ToOperator(String type2) {
        if (type2 == null) {
            return DataValidationConstraint.OperatorType.BETWEEN;
        }
        switch (type2) {
            case "bw":
                return DataValidationConstraint.OperatorType.BETWEEN;
            case "nb":
            case "nbw":
                return DataValidationConstraint.OperatorType.NOT_BETWEEN;
            case "eq":
                return DataValidationConstraint.OperatorType.EQUAL;
            case "ne":
                return DataValidationConstraint.OperatorType.NOT_EQUAL;
            case "gt":
            case "af":
                return DataValidationConstraint.OperatorType.GREATER_THAN;
            case "lt":
            case "bf":
                return DataValidationConstraint.OperatorType.LESS_THAN;
            case "gte":
            case "naf":
                return DataValidationConstraint.OperatorType.GREATER_OR_EQUAL;
            case "lte":
            case "nbf":
                return DataValidationConstraint.OperatorType.LESS_OR_EQUAL;
            default:
                return DataValidationConstraint.OperatorType.BETWEEN;
        }
    }

    private static String mapOperatorToLuckyType2(int op, boolean isDate) {
        switch (op) {
            case DataValidationConstraint.OperatorType.BETWEEN:
                return "bw";
            case DataValidationConstraint.OperatorType.NOT_BETWEEN:
                return "nb";
            case DataValidationConstraint.OperatorType.EQUAL:
                return "eq";
            case DataValidationConstraint.OperatorType.NOT_EQUAL:
                return "ne";
            case DataValidationConstraint.OperatorType.GREATER_THAN:
                return isDate ? "af" : "gt";
            case DataValidationConstraint.OperatorType.LESS_THAN:
                return isDate ? "bf" : "lt";
            case DataValidationConstraint.OperatorType.GREATER_OR_EQUAL:
                return isDate ? "naf" : "gte";
            case DataValidationConstraint.OperatorType.LESS_OR_EQUAL:
                return isDate ? "nbf" : "lte";
            default:
                return null;
        }
    }

    private static boolean isListType(String type) {
        return "dropdown".equalsIgnoreCase(type) || "checkbox".equalsIgnoreCase(type);
    }

    private static DataVerification copy(DataVerification src) {
        DataVerification dst = new DataVerification();
        dst.setType(src.getType());
        dst.setType2(src.getType2());
        dst.setValue1(src.getValue1());
        dst.setValue2(src.getValue2());
        dst.setChecked(src.getChecked());
        dst.setRemote(src.getRemote());
        dst.setProhibitInput(src.getProhibitInput());
        dst.setHintShow(src.getHintShow());
        dst.setHintText(src.getHintText());
        return dst;
    }

    private static int[] parseRowColKey(String key) {
        if (key == null) {
            return null;
        }
        int idx = key.indexOf('_');
        if (idx <= 0 || idx >= key.length() - 1) {
            return null;
        }
        try {
            return new int[]{Integer.parseInt(key.substring(0, idx)), Integer.parseInt(key.substring(idx + 1))};
        } catch (NumberFormatException e) {
            return null;
        }
    }
}
