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

import io.github.autoffice.luckysheet.model.cell.CellData;
import io.github.autoffice.luckysheet.model.cell.CellValue;
import io.github.autoffice.luckysheet.model.cell.Sparkline;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTExtension;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTExtensionList;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * 迷你图 (Sparkline) Luckysheet ↔ POI 双向映射器 (有损转换).
 *
 * <p>xlsx 的 sparkline 存储在 worksheet 的 extLst 扩展中. POI 无专用 API,
 * 采用直接读取 XML DOM 的方式导入; 导出时将 Luckysheet spl 信息写回扩展.</p>
 */
public final class SparklineMapper {

    private static final String EXT_URI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
    private static final String NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private static final String NS_XM = "http://schemas.microsoft.com/office/excel/2006/main";

    private SparklineMapper() {
    }

    /**
     * 从 Excel 工作表扩展中提取迷你图信息并写入 Luckysheet 单元格.
     *
     * @param sheet      源 POI 工作表
     * @param luckySheet 目标 Luckysheet 工作表
     */
    public static void mapToLuckySheet(XSSFSheet sheet, LuckySheet luckySheet) {
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        if (ctWorksheet == null || !ctWorksheet.isSetExtLst()) {
            return;
        }
        CTExtensionList extLst = ctWorksheet.getExtLst();
        for (CTExtension ext : extLst.getExtList()) {
            if (!EXT_URI.equalsIgnoreCase(ext.getUri())) {
                continue;
            }
            Node node = ext.getDomNode();
            NodeList groups = ((Element) node).getElementsByTagNameNS(NS_X14, "sparklineGroup");
            for (int i = 0; i < groups.getLength(); i++) {
                Element group = (Element) groups.item(i);
                String type = group.getAttribute("type");
                NodeList sparklines = group.getElementsByTagNameNS(NS_X14, "sparkline");
                for (int j = 0; j < sparklines.getLength(); j++) {
                    Element sl = (Element) sparklines.item(j);
                    String dataRange = textOfChild(sl, NS_XM, "f");
                    String target = textOfChild(sl, NS_XM, "sqref");
                    if (target == null) {
                        continue;
                    }
                    CellReference ref = new CellReference(target);
                    applyToCell(luckySheet, ref.getRow(), (int) ref.getCol(), type, dataRange);
                }
            }
        }
    }

    /**
     * 将 Luckysheet 单元格中的迷你图信息写入 Excel 工作表扩展.
     *
     * @param luckySheet 源 Luckysheet 工作表
     * @param sheet      目标 POI 工作表
     */
    public static void mapToExcel(LuckySheet luckySheet, XSSFSheet sheet) {
        if (luckySheet == null || CollectionUtils.isEmpty(luckySheet.getCelldata())) {
            return;
        }
        java.util.List<SparklineTarget> targets = new java.util.ArrayList<>();
        for (CellData cellData : luckySheet.getCelldata()) {
            CellValue value = cellData == null ? null : cellData.getV();
            Sparkline spl = value == null ? null : value.getSpl();
            if (spl == null || spl.getDataRange() == null) {
                continue;
            }
            targets.add(new SparklineTarget(cellData.getR(), cellData.getC(), spl));
        }
        if (targets.isEmpty()) {
            return;
        }

        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTExtensionList extLst = ctWorksheet.isSetExtLst() ? ctWorksheet.getExtLst() : ctWorksheet.addNewExtLst();
        CTExtension ext = findOrCreateExtension(extLst);

        Element extElement = (Element) ext.getDomNode();
        // 清理旧的 x14:sparklineGroups 以免累积
        NodeList oldGroups = extElement.getElementsByTagNameNS(NS_X14, "sparklineGroups");
        while (oldGroups.getLength() > 0) {
            extElement.removeChild(oldGroups.item(0));
        }
        Element groups = extElement.getOwnerDocument().createElementNS(NS_X14, "x14:sparklineGroups");
        groups.setAttribute("xmlns:xm", NS_XM);
        extElement.appendChild(groups);

        java.util.Map<String, Element> groupByType = new java.util.HashMap<>();
        for (SparklineTarget target : targets) {
            String type = target.spl.getType() == null ? "line" : target.spl.getType();
            Element group = groupByType.computeIfAbsent(type, t -> {
                Element g = extElement.getOwnerDocument().createElementNS(NS_X14, "x14:sparklineGroup");
                if (!"line".equalsIgnoreCase(t)) {
                    g.setAttribute("type", t);
                }
                Element sparklines = extElement.getOwnerDocument().createElementNS(NS_X14, "x14:sparklines");
                g.appendChild(sparklines);
                groups.appendChild(g);
                return g;
            });
            Element sparklines = (Element) group.getElementsByTagNameNS(NS_X14, "sparklines").item(0);
            Element sparkline = extElement.getOwnerDocument().createElementNS(NS_X14, "x14:sparkline");
            Element f = extElement.getOwnerDocument().createElementNS(NS_XM, "xm:f");
            f.appendChild(extElement.getOwnerDocument().createTextNode(target.spl.getDataRange()));
            Element sqref = extElement.getOwnerDocument().createElementNS(NS_XM, "xm:sqref");
            sqref.appendChild(extElement.getOwnerDocument().createTextNode(new CellReference(target.row, target.col).formatAsString()));
            sparkline.appendChild(f);
            sparkline.appendChild(sqref);
            sparklines.appendChild(sparkline);
        }
    }

    private static void applyToCell(LuckySheet luckySheet, int row, int col, String type, String dataRange) {
        CellData match = null;
        for (CellData cd : luckySheet.getCelldata()) {
            if (cd != null && cd.getR() != null && cd.getC() != null
                    && cd.getR() == row && cd.getC() == col) {
                match = cd;
                break;
            }
        }
        if (match == null) {
            match = LuckySheetFactory.createCellData();
            match.setR(row);
            match.setC(col);
            luckySheet.getCelldata().add(match);
        }
        if (match.getV() == null) {
            match.setV(new CellValue());
        }
        Sparkline spl = new Sparkline();
        spl.setType(type == null || type.isEmpty() ? "line" : type);
        spl.setDataRange(dataRange);
        match.getV().setSpl(spl);
    }

    private static String textOfChild(Element parent, String ns, String name) {
        NodeList list = parent.getElementsByTagNameNS(ns, name);
        if (list.getLength() == 0) {
            return null;
        }
        Node node = list.item(0).getFirstChild();
        return node == null ? null : node.getNodeValue();
    }

    private static CTExtension findOrCreateExtension(CTExtensionList extLst) {
        for (CTExtension ext : extLst.getExtList()) {
            if (EXT_URI.equalsIgnoreCase(ext.getUri())) {
                return ext;
            }
        }
        CTExtension ext = extLst.addNewExt();
        ext.setUri(EXT_URI);
        Element extElement = (Element) ext.getDomNode();
        extElement.setAttribute("xmlns:x14", NS_X14);
        return ext;
    }

    static final class SparklineTarget {
        final int row;
        final int col;
        final Sparkline spl;

        SparklineTarget(int row, int col, Sparkline spl) {
            this.row = row;
            this.col = col;
            this.spl = spl;
        }
    }
}
