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
package io.github.autoffice.luckysheet.model.cell;

import lombok.Data;

/**
 * 批注
 * <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/cell.html#%E5%9F%BA%E6%9C%AC%E5%8D%95%E5%85%83%E6%A0%BC"></a>
 *
 * @author hanbd
 */
@Data
public class Comment {
    /**
     * 批注框左边距
     */
    private Integer left;
    /**
     * 批注框上边距
     */
    private Integer top;
    /**
     * 批注框宽度
     */
    private Integer width;
    /**
     * 批注框高度
     */
    private Integer height;
    /**
     * 批注内容
     */
    private String value;
    /**
     * 批注框是否显示
     */
    private Boolean isshow;
}
