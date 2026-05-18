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

import io.github.autoffice.luckysheet.model.sheet.FilterColumn;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.model.sheet.Range;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

import java.util.Arrays;
import java.util.Map;

/**
 * 自动筛选 Luckysheet ↔ POI 双向映射器.
 *
 * <p>POI 精简 schema (poi-ooxml-lite) 中未包含 CTFilter/CTFilterColumn 类,
 * 因此具体的列筛选条件仅做透传, 保留 {@code filter_select} 范围. 这是本库相对
 * Luckysheet 完整筛选功能的有损点.</p>
 */
public final class FilterMapper {

    private FilterMapper() {
    }

    public static void mapToLuckySheet(XSSFSheet sheet, LuckySheet luckySheet) {
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        if (ctWorksheet == null || !ctWorksheet.isSetAutoFilter()) {
            return;
        }
        CTAutoFilter autoFilter = ctWorksheet.getAutoFilter();
        String ref = autoFilter.getRef();
        if (ref == null) {
            return;
        }
        CellRangeAddress cra = CellRangeAddress.valueOf(ref);
        Range range = new Range();
        range.setRow(Arrays.asList(cra.getFirstRow(), cra.getLastRow()));
        range.setColumn(Arrays.asList(cra.getFirstColumn(), cra.getLastColumn()));
        luckySheet.setFilter_select(range);
    }

    public static void mapToExcel(Range filterSelect, Map<String, FilterColumn> filters, XSSFSheet sheet) {
        if (filterSelect == null || CollectionUtils.isEmpty(filterSelect.getRow())
                || CollectionUtils.isEmpty(filterSelect.getColumn())
                || filterSelect.getRow().size() < 2 || filterSelect.getColumn().size() < 2) {
            return;
        }
        int firstRow = filterSelect.getRow().get(0);
        int lastRow = filterSelect.getRow().get(1);
        int firstCol = filterSelect.getColumn().get(0);
        int lastCol = filterSelect.getColumn().get(1);
        sheet.setAutoFilter(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        // filters 具体列条件因 poi-ooxml-lite 缺失 CTFilterColumn 等类而无法写入,
        // 这里仅保留筛选范围. Luckysheet 再次加载时会显示筛选按钮但各列条件为默认.
    }
}
