package io.github.autoffice.luckysheet.model.sheet;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#frozen"">
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
