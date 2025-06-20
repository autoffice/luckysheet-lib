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

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;

@Data
public class SheetImage {
    /**
     * 图片定位类型
     */
    private ImageType type;

    /**
     * 图片url信息，可以是base64格式的图片
     */
    private String src;

    /**
     * 图片原始宽度
     */
    private Integer originWidth;

    /**
     * 图片原始高度
     */
    private Integer originHeight;

    /**
     * 图片位置信息
     */
    @JsonProperty("default")
    private ImagePosition position;

    /**
     * 图片裁剪信息
     */
    private ImageCrop crop;

    /**
     * 固定位置
     */
    @JsonProperty("isFixedPos")
    private boolean isFixedPos;

    /**
     * 固定位置左位移
     */
    private Integer fixedLeft;
    /**
     * /固定位置顶位移
     */
    private Integer fixedTop;

    /**
     * 问题
     */
    private ImageBorder border;
}
