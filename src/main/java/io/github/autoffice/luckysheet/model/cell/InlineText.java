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

@Data
public class InlineText {
    /**
     * 字体
     */
    private FontFamily ff;
    /**
     * 字体颜色. #fff000
     */
    private String fc;
    /**
     * 字体大小
     */
    private Short fs;

    /**
     * 删除线. 0 常规 、 1 删除线
     */
    private Cancelline cl;
    /**
     * 下划线. 0 无 、 1 有
     */
    private Underline un;
    /**
     * 粗体. 0 常规 、 1加粗
     */
    private Bold bl;
    /**
     * 斜体. 0 常规 、 1 斜体
     */
    private Italic it;

    /**
     * 文本内容
     */
    private String v;
}
