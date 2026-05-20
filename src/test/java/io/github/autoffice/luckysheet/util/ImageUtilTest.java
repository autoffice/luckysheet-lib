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
import org.junit.jupiter.api.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Base64;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 测试 {@link ImageUtil} 的图片处理功能, 涵盖 Base64 编解码、PNG 转换以及 ID 生成.
 */
class ImageUtilTest {

    // ========== toLuckySheetImage ==========

    @Test
    void toLuckySheetImage_nullData_throwsNpe() {
        // ByteArrayInputStream 构造函数不接受 null, 会抛出 NPE
        assertThrows(NullPointerException.class, () -> ImageUtil.toLuckySheetImage(null));
    }

    @Test
    void toLuckySheetImage_emptyData_returnsNull() {
        // 空数组传入 ImageIO.read 返回 null, 内部 toPng 返回 null
        String result = ImageUtil.toLuckySheetImage(new byte[0]);
        assertNull(result);
    }

    @Test
    void toLuckySheetImage_validPngBytes_returnsBase64DataUri() throws IOException {
        byte[] pngBytes = createPngBytes(10, 10);
        String result = ImageUtil.toLuckySheetImage(pngBytes);
        assertNotNull(result);
        assertTrue(result.startsWith("data:image/png;base64,"));
    }

    @Test
    void toLuckySheetImage_invalidBytes_returnsNull() {
        // 随机字节, 无法识别为图片格式, ImageIO.read 返回 null
        byte[] randomBytes = new byte[]{1, 2, 3, 4, 5, 6, 7, 8};
        String result = ImageUtil.toLuckySheetImage(randomBytes);
        assertNull(result);
    }

    // ========== getImageData ==========

    @Test
    void getImageData_nullImage_throwsNpe() {
        // 实现未对 null image 做防护, 会触发 NPE
        assertThrows(NullPointerException.class, () -> ImageUtil.getImageData(null));
    }

    @Test
    void getImageData_srcIsNull_returnsNull() {
        SheetImage image = new SheetImage();
        image.setSrc(null);
        byte[] result = ImageUtil.getImageData(image);
        assertNull(result);
    }

    @Test
    void getImageData_validBase64DataUri_returnsBytes() {
        byte[] expected = new byte[]{1, 2, 3, 4, 5};
        String base64 = Base64.getEncoder().encodeToString(expected);
        SheetImage image = new SheetImage();
        image.setSrc("data:image/png;base64," + base64);
        byte[] result = ImageUtil.getImageData(image);
        assertNotNull(result);
        assertArrayEquals(expected, result);
    }

    @Test
    void getImageData_srcMissingDataPrefix_returnsNull() {
        // 普通 URL 不含 base64, 标记, 应返回 null
        SheetImage image = new SheetImage();
        image.setSrc("https://example.com/image.png");
        byte[] result = ImageUtil.getImageData(image);
        assertNull(result);
    }

    // ========== getLuckySheetImageId ==========

    @Test
    void getLuckySheetImageId_returnsNonNullPrefixedId() {
        String id = ImageUtil.getLuckySheetImageId();
        assertNotNull(id);
        assertTrue(id.startsWith("img_"));
    }

    @Test
    void getLuckySheetImageId_callTwice_bothNonNullAndPrefixed() {
        String id1 = ImageUtil.getLuckySheetImageId();
        String id2 = ImageUtil.getLuckySheetImageId();
        assertNotNull(id1);
        assertNotNull(id2);
        assertTrue(id1.startsWith("img_"));
        assertTrue(id2.startsWith("img_"));
    }

    // ========== helpers ==========

    private byte[] createPngBytes(int w, int h) throws IOException {
        BufferedImage img = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(img, "PNG", baos);
        return baos.toByteArray();
    }
}
