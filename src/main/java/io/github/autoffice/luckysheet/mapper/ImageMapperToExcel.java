package io.github.autoffice.luckysheet.mapper;

import io.github.autoffice.luckysheet.model.image.SheetImage;
import io.github.autoffice.luckysheet.util.ImageUtil;
import io.github.autoffice.luckysheet.util.NumberUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.MapUtils;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Map;

@Slf4j
public class ImageMapperToExcel {

    public static void mapToSheet(Map<String, SheetImage> images, XSSFSheet sheet) {
        if (MapUtils.isEmpty(images)) {
            return;
        }

        for (SheetImage image : images.values()) {
            mapToSheet(image, sheet);
        }
    }

    private static void mapToSheet(SheetImage image, XSSFSheet sheet) {
        // luckysheet only support png base64 image
        int index = sheet.getWorkbook().addPicture(ImageUtil.getImageData(image), Workbook.PICTURE_TYPE_PNG);
        ClientAnchor clientAnchor = getClientAnchor(image, sheet);
        sheet.createDrawingPatriarch().createPicture(clientAnchor, index);
    }

    private static ClientAnchor getClientAnchor(SheetImage image, XSSFSheet sheet) {
        CreationHelper creationHelper = sheet.getWorkbook().getCreationHelper();
        ClientAnchor clientAnchor = creationHelper.createClientAnchor();

        int sumRowHeightPixel = 0;
        int row1;
        int dy1;
        for (int i = 0; ; i++) {
            XSSFRow row = sheet.getRow(i);
            float rowHeightInPoints;
            if (row == null) {
                rowHeightInPoints = sheet.getDefaultRowHeightInPoints();
            } else {
                rowHeightInPoints = row.getHeightInPoints();
            }

            sumRowHeightPixel += Units.pointsToPixel(rowHeightInPoints);
            if (image.getPosition().getTop() <= sumRowHeightPixel) {
                row1 = i;
                dy1 = Units.pointsToPixel(rowHeightInPoints) - (sumRowHeightPixel - image.getPosition().getTop());
                break;
            }
        }

        sumRowHeightPixel = 0;
        int row2;
        int dy2;
        for (int i = 0; ; i++) {
            XSSFRow row = sheet.getRow(i);
            float rowHeightInPoints;
            if (row == null) {
                rowHeightInPoints = sheet.getDefaultRowHeightInPoints();
            } else {
                rowHeightInPoints = row.getHeightInPoints();
            }

            sumRowHeightPixel += Units.pointsToPixel(rowHeightInPoints);
            if (image.getPosition().getTop() + image.getPosition().getHeight() <= sumRowHeightPixel) {
                row2 = i;
                dy2 = Units.pointsToPixel(rowHeightInPoints) - (sumRowHeightPixel - image.getPosition().getTop() - image.getPosition().getHeight());
                break;
            }
        }

        float sumColWidthPixel = 0;
        int col1;
        float dx1;
        for (int i = 0; ; i++) {
            sumColWidthPixel += sheet.getColumnWidthInPixels(i);
            if (image.getPosition().getLeft() <= sumColWidthPixel) {
                col1 = i;
                dx1 = sheet.getColumnWidthInPixels(i) - (sumColWidthPixel - image.getPosition().getLeft());
                break;
            }
        }

        sumColWidthPixel = 0;
        int col2;
        float dx2;
        for (int i = 0; ; i++) {
            sumColWidthPixel += sheet.getColumnWidthInPixels(i);
            if (image.getPosition().getLeft() + image.getPosition().getWidth() <= sumColWidthPixel) {
                col2 = i;
                dx2 = sheet.getColumnWidthInPixels(i) - (sumColWidthPixel - image.getPosition().getLeft() - image.getPosition().getWidth());
                break;
            }
        }

        clientAnchor.setRow1(row1);
        clientAnchor.setCol1(col1);
        clientAnchor.setRow2(row2);
        clientAnchor.setCol2(col2);
        clientAnchor.setDx1(NumberUtil.pixelToEMU(dx1));
        clientAnchor.setDx2(NumberUtil.pixelToEMU(dx2));
        clientAnchor.setDy1(Units.pixelToEMU(dy1));
        clientAnchor.setDy2(Units.pixelToEMU(dy2));

        if (image.getType() != null) {
            clientAnchor.setAnchorType(image.getType().getPoiValue());
        }

        return clientAnchor;
    }
}
