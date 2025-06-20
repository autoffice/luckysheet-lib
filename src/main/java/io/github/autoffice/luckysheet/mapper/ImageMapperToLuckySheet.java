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

import io.github.autoffice.luckysheet.model.image.ImagePosition;
import io.github.autoffice.luckysheet.model.image.ImageType;
import io.github.autoffice.luckysheet.model.image.SheetImage;
import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import io.github.autoffice.luckysheet.util.ImageUtil;
import io.github.autoffice.luckysheet.util.NumberUtil;
import org.apache.poi.ss.util.ImageUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.awt.*;
import java.util.List;

public class ImageMapperToLuckySheet {
    public static void mapToSheet(XSSFSheet sheet, LuckySheet luckySheet) {
        XSSFDrawing drawingPatriarch = sheet.getDrawingPatriarch();
        if (drawingPatriarch == null) {
            return;
        }

        List<XSSFShape> shapes = drawingPatriarch.getShapes();
        for (XSSFShape shape : shapes) {
            if (shape instanceof XSSFPicture) {
                XSSFPicture picture = (XSSFPicture) shape;
                String image = ImageUtil.toLuckySheetImage(picture.getPictureData().getData());
                SheetImage sheetImage = LuckySheetFactory.createImage();
                sheetImage.setSrc(image);
                sheetImage.setFixedPos(false);

                ImagePosition imagePosition = createImagePosition(sheet, picture);
                sheetImage.setPosition(imagePosition);
                sheetImage.setOriginWidth(imagePosition.getWidth());
                sheetImage.setOriginHeight(imagePosition.getHeight());
                sheetImage.getCrop().setWidth(imagePosition.getWidth());
                sheetImage.getCrop().setHeight(imagePosition.getHeight());
                sheetImage.getCrop().setOffsetLeft(0);
                sheetImage.getCrop().setOffsetTop(0);

                sheetImage.setType(ImageType.of(picture.getClientAnchor().getAnchorType()));
                luckySheet.getImages().put(ImageUtil.getLuckySheetImageId(), sheetImage);
            }
        }
    }

    private static ImagePosition createImagePosition(XSSFSheet sheet, XSSFPicture picture) {
        Dimension dimensionFromAnchor = ImageUtils.getDimensionFromAnchor(picture);
        XSSFClientAnchor clientAnchor = picture.getClientAnchor();

        int topPixel = 0;
        for (int i = 0; i < clientAnchor.getRow1(); i++) {
            XSSFRow row = sheet.getRow(i);
            float rowHeightInPoints;
            if (row == null) {
                rowHeightInPoints = sheet.getDefaultRowHeightInPoints();
            } else {
                rowHeightInPoints = row.getHeightInPoints();
            }

            topPixel += Units.pointsToPixel(rowHeightInPoints);
        }

        topPixel += NumberUtil.emu2Pixel(clientAnchor.getDy1());

        float leftPixel = 0;
        for (int i = 0; i < clientAnchor.getCol1(); i++) {
            leftPixel += sheet.getColumnWidthInPixels(i);
        }

        leftPixel += NumberUtil.emu2Pixel(clientAnchor.getDx1());

        ImagePosition imagePosition = new ImagePosition();
        imagePosition.setTop(topPixel);
        imagePosition.setLeft(Math.round(leftPixel));
        imagePosition.setWidth(NumberUtil.emu2Pixel(dimensionFromAnchor.getWidth()));
        imagePosition.setHeight(NumberUtil.emu2Pixel(dimensionFromAnchor.getHeight()));

        return imagePosition;
    }
}
