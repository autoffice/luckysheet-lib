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
package io.github.autoffice.luckysheet.model.image;

import lombok.Data;

/**
 * 图片边框配置, 包含边框宽度、样式和颜色.
 *
 * <p>对应 Luckysheet 图片对象中的 border 属性.</p>
 *
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#images">Luckysheet 图片文档</a>
 */
@Data
public class ImageBorder {

    /**
     * 边框宽度
     */
    private Integer width;

    /**
     * 边框样式类型
     */
    private ImageBorderType style;

    /**
     * 边框验收
     */
    private String color;
}
