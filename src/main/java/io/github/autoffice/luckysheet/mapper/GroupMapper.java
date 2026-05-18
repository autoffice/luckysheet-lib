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

import io.github.autoffice.luckysheet.model.sheet.Group;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

import java.util.ArrayList;
import java.util.List;

/**
 * 行列分组 (大纲) Luckysheet ↔ POI 双向映射器.
 */
public final class GroupMapper {

    private GroupMapper() {
    }

    public static void mapToLuckySheet(XSSFSheet sheet, LuckySheet luckySheet) {
        List<Group> rowGroups = collectRowGroups(sheet);
        if (!rowGroups.isEmpty()) {
            luckySheet.setRowGroup(rowGroups);
        }

        List<Group> colGroups = collectColumnGroups(sheet);
        if (!colGroups.isEmpty()) {
            luckySheet.setColGroup(colGroups);
        }
    }

    public static void mapToExcel(List<Group> rowGroups, List<Group> colGroups, XSSFSheet sheet) {
        if (CollectionUtils.isNotEmpty(rowGroups)) {
            for (Group g : rowGroups) {
                if (g == null || g.getStart() == null || g.getEnd() == null) {
                    continue;
                }
                sheet.groupRow(g.getStart(), g.getEnd());
                if (Boolean.TRUE.equals(g.getCollapsed())) {
                    sheet.setRowGroupCollapsed(g.getStart(), true);
                }
            }
        }
        if (CollectionUtils.isNotEmpty(colGroups)) {
            for (Group g : colGroups) {
                if (g == null || g.getStart() == null || g.getEnd() == null) {
                    continue;
                }
                sheet.groupColumn(g.getStart(), g.getEnd());
                if (Boolean.TRUE.equals(g.getCollapsed())) {
                    sheet.setColumnGroupCollapsed(g.getStart(), true);
                }
            }
        }
    }

    private static List<Group> collectRowGroups(XSSFSheet sheet) {
        List<Group> groups = new ArrayList<>();
        int first = sheet.getFirstRowNum();
        int last = sheet.getLastRowNum();
        Group current = null;
        for (int i = first; i <= last; i++) {
            XSSFRow row = sheet.getRow(i);
            int level = row == null ? 0 : row.getCTRow().getOutlineLevel();
            boolean hidden = row != null && row.getZeroHeight();
            if (level <= 0) {
                if (current != null) {
                    groups.add(current);
                    current = null;
                }
                continue;
            }
            if (current == null || !current.getLevel().equals(level)) {
                if (current != null) {
                    groups.add(current);
                }
                current = new Group();
                current.setStart(i);
                current.setEnd(i);
                current.setLevel(level);
                current.setCollapsed(hidden);
            } else {
                current.setEnd(i);
                if (!Boolean.TRUE.equals(current.getCollapsed()) && hidden) {
                    current.setCollapsed(true);
                }
            }
        }
        if (current != null) {
            groups.add(current);
        }
        return groups;
    }

    private static List<Group> collectColumnGroups(XSSFSheet sheet) {
        List<Group> groups = new ArrayList<>();
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        if (ctWorksheet == null) {
            return groups;
        }
        List<CTCols> colsList = ctWorksheet.getColsList();
        if (colsList == null) {
            return groups;
        }
        for (CTCols cols : colsList) {
            List<CTCol> colList = cols.getColList();
            if (colList == null) {
                continue;
            }
            for (CTCol col : colList) {
                int level = col.getOutlineLevel();
                if (level <= 0) {
                    continue;
                }
                Group g = new Group();
                g.setStart((int) (col.getMin() - 1));
                g.setEnd((int) (col.getMax() - 1));
                g.setLevel(level);
                g.setCollapsed(col.getHidden());
                groups.add(g);
            }
        }
        return groups;
    }
}
