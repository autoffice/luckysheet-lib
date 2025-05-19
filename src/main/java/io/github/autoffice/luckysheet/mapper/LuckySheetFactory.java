package io.github.autoffice.luckysheet.mapper;

import io.github.autoffice.luckysheet.model.LuckyFile;
import io.github.autoffice.luckysheet.model.LuckyFileInfo;
import io.github.autoffice.luckysheet.model.cell.CellData;
import io.github.autoffice.luckysheet.model.cell.CellType;
import io.github.autoffice.luckysheet.model.cell.CellValue;
import io.github.autoffice.luckysheet.model.cell.Comment;
import io.github.autoffice.luckysheet.model.cell.InlineText;
import io.github.autoffice.luckysheet.model.image.ImageBorder;
import io.github.autoffice.luckysheet.model.image.ImageCrop;
import io.github.autoffice.luckysheet.model.image.ImagePosition;
import io.github.autoffice.luckysheet.model.image.SheetImage;
import io.github.autoffice.luckysheet.model.sheet.Border;
import io.github.autoffice.luckysheet.model.sheet.BorderRangeType;
import io.github.autoffice.luckysheet.model.sheet.BorderStyleType;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.util.NumberUtil;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class LuckySheetFactory {
    public static LuckyFile createLuckyFile() {
        return new LuckyFile();
    }

    public static LuckyFileInfo createLuckyFileInfo() {
        return new LuckyFileInfo();
    }

    public static LuckySheet createLuckySheet(String sheetName) {
        LuckySheet luckySheet = new LuckySheet();
        luckySheet.setName(sheetName);
        return luckySheet;
    }

    public static CellData createCellData() {
        CellData cellData = new CellData();
        CellValue v = new CellValue();
        v.setCt(new CellType());
        cellData.setV(v);

        return cellData;
    }

    public static InlineText createInlineText() {
        return new InlineText();
    }

    public static Comment createComment() {
        return new Comment();
    }

    public static Border createBorder() {
        Border border = new Border();
        border.setRangeType(BorderRangeType.CELL);
        border.setValue(new Border.Value());
        return border;
    }

    public static Border.Style createBorderStyle(XSSFColor borderColor, BorderStyle borderStyle) {
        if (borderColor == null) {
            return null;
        }

        Border.Style style = new Border.Style();
        style.setColor(NumberUtil.rgbToColorString(borderColor.getRGB()));
        style.setStyle(BorderStyleType.of(borderStyle));
        return style;
    }

    public static SheetImage createImage() {
        SheetImage sheetImage = new SheetImage();
        sheetImage.setBorder(new ImageBorder());
        sheetImage.setCrop(new ImageCrop());
        sheetImage.setPosition(new ImagePosition());
        return sheetImage;
    }

    public static boolean hasBorderStyle(Border border) {
        return border.getValue().getT() != null
                || border.getValue().getR() != null
                || border.getValue().getB() != null
                || border.getValue().getL() != null;
    }

    public static Border createBorder(int row, int col) {
        Border borderTmp = new Border();
        borderTmp.setRangeType(BorderRangeType.CELL);
        Border.Value value = new Border.Value();
        value.setRow_index(row);
        value.setCol_index(col);
        borderTmp.setValue(value);
        return borderTmp;
    }

    public static Border.Style createBorderStyle(Border border) {
        Border.Style style = new Border.Style();
        style.setStyle(border.getStyle());
        style.setColor(border.getColor());
        return style;
    }
}
