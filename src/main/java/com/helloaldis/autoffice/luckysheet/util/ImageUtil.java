package com.helloaldis.autoffice.luckysheet.util;

import com.helloaldis.autoffice.luckysheet.model.image.SheetImage;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.commons.lang3.StringUtils;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Base64;

@Slf4j
public class ImageUtil {
    public static String toLuckySheetImage(byte[] data) {
        byte[] png = toPng(data);
        if (png == null) {
            return null;
        }

        String base64 = Base64.getEncoder().encodeToString(data);
        return "data:image/png;base64," + base64;
    }

    private static byte[] toPng(byte[] data) {
        ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(data);
        try {
            BufferedImage image = ImageIO.read(byteArrayInputStream);
            if (image == null) {
                throw new IOException("图片解析为null");
            }

            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            ImageIO.write(image, "PNG", byteArrayOutputStream);
            return byteArrayOutputStream.toByteArray();
        } catch (Exception e) {
            log.error("图片转png失败", e);
        }

        return null;
    }

    public static byte[] getImageData(SheetImage image) {
        if (isBase64Imgae(image)) {
            String[] split = image.getSrc().split(",");
            return Base64.getDecoder().decode(split[1]);
        }

        return null;
    }

    private static boolean isBase64Imgae(SheetImage image) {
        return StringUtils.contains(image.getSrc(), "base64,");
    }

    public static String getLuckySheetImageId() {
        String time = String.valueOf(System.currentTimeMillis());
        String random = RandomStringUtils.randomAlphabetic(4);
        String substring = time.substring(time.length()-8);
        return "img_" + substring + random + "_" + time;
    }
}
