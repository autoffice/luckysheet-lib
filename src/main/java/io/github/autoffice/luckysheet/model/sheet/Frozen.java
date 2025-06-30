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
package io.github.autoffice.luckysheet.model.sheet;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#frozen">
 *  冻结行列设置，分为6种类型 </a>
 */
@Data
public class Frozen {
    private FrozenType type;
    private Range range;

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public static class Range{
        private int row_focus;
        private int column_focus;
    }
}
