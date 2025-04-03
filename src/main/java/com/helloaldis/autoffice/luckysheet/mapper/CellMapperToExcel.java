package com.helloaldis.autoffice.luckysheet.mapper;

import com.helloaldis.autoffice.luckysheet.model.cell.Bold;
import com.helloaldis.autoffice.luckysheet.model.cell.Cancelline;
import com.helloaldis.autoffice.luckysheet.model.cell.CellData;
import com.helloaldis.autoffice.luckysheet.model.cell.CellHorizontalType;
import com.helloaldis.autoffice.luckysheet.model.cell.CellType;
import com.helloaldis.autoffice.luckysheet.model.cell.CellTypeEnum;
import com.helloaldis.autoffice.luckysheet.model.cell.CellValue;
import com.helloaldis.autoffice.luckysheet.model.cell.CellVerticalType;
import com.helloaldis.autoffice.luckysheet.model.cell.Comment;
import com.helloaldis.autoffice.luckysheet.model.cell.FontFamily;
import com.helloaldis.autoffice.luckysheet.model.cell.InlineText;
import com.helloaldis.autoffice.luckysheet.model.cell.Italic;
import com.helloaldis.autoffice.luckysheet.model.cell.TextBreakType;
import com.helloaldis.autoffice.luckysheet.model.cell.TextRotateType;
import com.helloaldis.autoffice.luckysheet.model.cell.Underline;
import com.helloaldis.autoffice.luckysheet.util.NumberUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.util.List;

@Slf4j
public class CellMapperToExcel {
    public static void mapToCell(CellData cellData, XSSFCell cell) {
        mapCellValue(cellData, cell);
        mapCellStyle(cellData, cell);
        mapFunction(cellData.getV().getF(), cell);
        mapComment(cellData.getV().getPs(), cell);
    }

    private static void mapComment(Comment ps, XSSFCell cell) {
        if (ps == null) {
            return;
        }

        org.apache.poi.ss.usermodel.Comment comment = PoiFactory.createComment(cell);
        if (StringUtils.isNotEmpty(ps.getValue())) {
            XSSFRichTextString xssfRichTextString = PoiFactory.createRichTextString(cell);
            xssfRichTextString.append(ps.getValue());
            comment.setString(xssfRichTextString);
        }

        comment.setAddress(cell.getAddress());
        comment.setVisible(ps.getIsshow());
        cell.setCellComment(comment);
    }

    private static void mapFunction(String func, XSSFCell cell) {
        if (StringUtils.isEmpty(func)) {
            return;
        }

        try {
            if (func.startsWith("=")) {
                func = func.substring(1);
            }

            cell.setCellFormula(func);
        } catch (Exception e) {
            log.error("func {} not supported", func, e);
        }
    }

    private static void mapCellValue(CellData cellData, XSSFCell cell) {
        if (cellData.getV().getV() == null) {
            if (cellData.getV().getCt() != null &&
                    cellData.getV().getCt().getT() == CellTypeEnum.INLINESTR) {
                mapRichTextString(cellData.getV().getCt().getS(), cell);
            }
            return;
        }

        if (cellData.getV().getV() instanceof Boolean) {
            cell.setCellValue((Boolean) cellData.getV().getV());
        } else if (cellData.getV().getV() instanceof Integer) {
            if (cellData.getV().getCt() != null && cellData.getV().getCt().getT() != null && cellData.getV().getCt().getT() == CellTypeEnum.DATETIME) {
                // 日期格式数据使用monitor值
                cell.setCellValue(cellData.getV().getM());
            } else {
                cell.setCellValue(Double.valueOf((Integer) cellData.getV().getV()));
            }
        } else if (cellData.getV().getV() instanceof Double) {
            cell.setCellValue((Double) cellData.getV().getV());
        } else if (cellData.getV().getV() instanceof String) {
            cell.setCellValue((String) cellData.getV().getV());
        }
    }

    private static void mapRichTextString(List<InlineText> s, XSSFCell cell) {
        if (CollectionUtils.isEmpty(s)) {
            return;
        }

        XSSFRichTextString xssfRichTextString = PoiFactory.createRichTextString(cell);
        for (InlineText inlineText : s) {
            if (inlineText.getV() != null) {
                XSSFFont font = PoiFactory.createFont(cell);
                mapFont(font, inlineText.getFf(),
                        inlineText.getBl(),
                        inlineText.getIt(),
                        inlineText.getUn(),
                        inlineText.getCl(),
                        inlineText.getFs(),
                        inlineText.getFc());
                xssfRichTextString.append(inlineText.getV(), font);
            }
        }
        cell.setCellValue(xssfRichTextString);
    }


    private static void mapCellStyle(CellData cellData, XSSFCell cell) {
        XSSFCellStyle xssfCellStyle = PoiFactory.createCellStyle(cell);

        mapBackgroundColor(cellData.getV().getBg(), xssfCellStyle);
        mapFont(cellData.getV(), cell, xssfCellStyle);
        mapHorizontalAlignment(cellData.getV().getHt(), xssfCellStyle);
        mapVerticalAlignment(cellData.getV().getVt(), xssfCellStyle);
        mapRotation(cellData.getV().getTr(), xssfCellStyle);
        mapDataFormat(cellData.getV().getCt(), xssfCellStyle);
        mapWrapText(cellData.getV().getTb(), xssfCellStyle);

        cell.setCellStyle(xssfCellStyle);
    }

    private static void mapWrapText(TextBreakType tb, XSSFCellStyle xssfCellStyle) {
        if (tb == null) {
            return;
        }

        if (tb == TextBreakType.LINE_WRAP) {
            xssfCellStyle.setWrapText(true);
        }
    }

    private static void mapDataFormat(CellType ct, XSSFCellStyle xssfCellStyle) {
        if (ct == null || ct.getFa() == null) {
            return;
        }

        int builtinFormat = BuiltinFormats.getBuiltinFormat(ct.getFa());

        if (builtinFormat != -1) {
            xssfCellStyle.setDataFormat(builtinFormat);
        }
    }

    private static void mapBackgroundColor(String bg, XSSFCellStyle xssfCellStyle) {
        if (StringUtils.isEmpty(bg)) {
            return;
        }

        byte[] rgb = NumberUtil.colorStringToRgb(bg);
        XSSFColor xssfColor = new XSSFColor(rgb, null);
        xssfCellStyle.setFillForegroundColor(xssfColor);
        xssfCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    private static void mapFont(CellValue cellValue, XSSFCell cell, XSSFCellStyle xssfCellStyle) {
        // 富文本无需重复设置字体样式
        if (cellValue.getCt() != null && cellValue.getCt().getT() == CellTypeEnum.INLINESTR) {
            return;
        }

        XSSFFont font = PoiFactory.createFont(cell);

        mapFont(font, cellValue.getFf(),
                cellValue.getBl(),
                cellValue.getIt(),
                cellValue.getUn(),
                cellValue.getCl(),
                cellValue.getFs(),
                cellValue.getFc());

        xssfCellStyle.setFont(font);
    }

    private static void mapFont(XSSFFont font, FontFamily ff, Bold bl, Italic it,
                                Underline un, Cancelline cl, Short fs, String fc) {
        if (ff != null) {
            font.setFontName(ff.getPoiValue());
        }

        if (bl != null) {
            font.setBold(bl.isPoiValue());
        }

        if (it != null) {
            font.setItalic(it.isPoiValue());
        }

        if (un != null) {
            font.setUnderline(un.getPoiValue());
        }

        if (cl != null) {
            font.setStrikeout(cl.isPoiValue());
        }

        if (fs != null) {
            font.setFontHeightInPoints(fs);
        }

        if (StringUtils.isNotEmpty(fc)) {
            byte[] rgb = NumberUtil.colorStringToRgb(fc);
            XSSFColor xssfColor = new XSSFColor(rgb, null);
            font.setColor(xssfColor);
        }
    }


    public static void mapHorizontalAlignment(CellHorizontalType cellHorizontalType, XSSFCellStyle xssfCellStyle) {
        if (cellHorizontalType == null) {
            return;
        }

        xssfCellStyle.setAlignment(cellHorizontalType.getPoiValue());
    }

    public static void mapVerticalAlignment(CellVerticalType cellHorizontalType, XSSFCellStyle xssfCellStyle) {
        if (cellHorizontalType == null) {
            return;
        }

        xssfCellStyle.setVerticalAlignment(cellHorizontalType.getPoiValue());
    }

    private static void mapRotation(TextRotateType textrotate, XSSFCellStyle xssfCellStyle) {
        if (textrotate == null) {
            return;
        }

        xssfCellStyle.setRotation(textrotate.getPoiValue());
    }
}
