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

import io.github.autoffice.luckysheet.model.sheet.Hyperlink;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 超链接 Luckysheet ↔ POI 双向映射器.
 */
public final class HyperlinkMapper {

    private HyperlinkMapper() {
    }

    public static void mapToLuckySheet(XSSFSheet sheet, LuckySheet luckySheet) {
        List<XSSFHyperlink> hyperlinks = sheet.getHyperlinkList();
        if (hyperlinks == null || hyperlinks.isEmpty()) {
            return;
        }

        Map<String, Hyperlink> map = luckySheet.getHyperlink();
        if (map == null) {
            map = new HashMap<>();
            luckySheet.setHyperlink(map);
        }

        for (XSSFHyperlink hyperlink : hyperlinks) {
            int row = hyperlink.getFirstRow();
            int col = hyperlink.getFirstColumn();
            Hyperlink model = new Hyperlink();
            model.setLinkAddress(hyperlink.getAddress());
            model.setLinkTooltip(hyperlink.getTooltip());
            HyperlinkType type = hyperlink.getType();
            if (type == HyperlinkType.DOCUMENT) {
                model.setLinkType("internal");
            } else {
                model.setLinkType("external");
            }
            map.put(row + "_" + col, model);
        }
    }

    public static void mapToExcel(Map<String, Hyperlink> hyperlinks, XSSFSheet sheet) {
        if (MapUtils.isEmpty(hyperlinks)) {
            return;
        }

        XSSFCreationHelper helper = sheet.getWorkbook().getCreationHelper();
        for (Map.Entry<String, Hyperlink> entry : hyperlinks.entrySet()) {
            String key = entry.getKey();
            Hyperlink value = entry.getValue();
            if (value == null || StringUtils.isBlank(value.getLinkAddress())) {
                continue;
            }
            int[] rc = parseRowColKey(key);
            if (rc == null) {
                continue;
            }
            HyperlinkType type = resolveType(value.getLinkType(), value.getLinkAddress());
            XSSFHyperlink link = helper.createHyperlink(type);
            link.setAddress(value.getLinkAddress());
            if (StringUtils.isNotBlank(value.getLinkTooltip())) {
                link.setTooltip(value.getLinkTooltip());
            }
            link.setCellReference(new CellReference(rc[0], rc[1]).formatAsString());
            PoiFactory.createOrGetCell(sheet, rc[0], rc[1]).setHyperlink(link);
        }
    }

    private static HyperlinkType resolveType(String linkType, String address) {
        if ("internal".equalsIgnoreCase(linkType)) {
            return HyperlinkType.DOCUMENT;
        }
        if (address != null && address.toLowerCase().startsWith("mailto:")) {
            return HyperlinkType.EMAIL;
        }
        return HyperlinkType.URL;
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
            int r = Integer.parseInt(key.substring(0, idx));
            int c = Integer.parseInt(key.substring(idx + 1));
            return new int[]{r, c};
        } catch (NumberFormatException e) {
            return null;
        }
    }
}
