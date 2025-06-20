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

@Data
public class ImageCrop {
    /**
     * 图片裁剪后宽度
     */
    private Integer width;
    /**
     * 图片裁剪后高度
     */
    private Integer height;
    /**
     * 图片裁剪后离未裁剪时左边的位移
     */
    private Integer offsetLeft;
    /**
     * 图片裁剪后离未裁剪时顶部的位移
     */
    private Integer offsetTop;
}
