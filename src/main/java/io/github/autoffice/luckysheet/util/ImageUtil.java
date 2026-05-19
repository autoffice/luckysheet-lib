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
package io.github.autoffice.luckysheet.util;

import io.github.autoffice.luckysheet.model.image.SheetImage;
import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.commons.lang3.StringUtils;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Base64;

/**
 * 图片处理工具类.
 *
 * <p>提供图片格式转换、Base64 编解码和 Luckysheet 图片 ID 生成等功能.</p>
 */
@NoArgsConstructor(access = AccessLevel.PRIVATE)
@Slf4j
public final class ImageUtil {
    /**
     * 将图片字节数组转换为 Base64 Data URI 字符串.
     *
     * @param data 图片字节数组
     * @return Base64 Data URI 字符串, 转换失败时返回 null
     */
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

    /**
     * 从 SheetImage 中提取图片字节数据.
     *
     * @param image Luckysheet 图片对象
     * @return 图片字节数组, 非 Base64 图片时返回 null
     */
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

    /**
     * 生成 Luckysheet 格式的图片唯一标识符.
     *
     * @return 以 "img_" 为前缀的唯一图片 ID
     */
    public static String getLuckySheetImageId() {
        String time = String.valueOf(System.currentTimeMillis());
        String random = RandomStringUtils.randomAlphabetic(4);
        String substring = time.substring(time.length()-8);
        return "img_" + substring + random + "_" + time;
    }
}
