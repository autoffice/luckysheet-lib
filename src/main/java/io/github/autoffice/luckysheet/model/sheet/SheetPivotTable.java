package io.github.autoffice.luckysheet.model.sheet;

import lombok.Data;

import java.util.List;

/**
 * 数据透视表设置
 * @see <a href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#pivottable">luckysheet官方文档</a>
 */
@Data
public class SheetPivotTable {

    private Range pivot_select_save;
    /**
     * 源数据所在的sheet页
     */
    private Integer pivotDataSheetIndex;
    private List<PivotColRow> column;
    private List<PivotColRow> row;

    // TODO
    private List<Object> filter;

    private List<PivotValue> values;

    private String showType;

    /**
     * 数据透视表的源数据
     */
    private List<List<Object>> pivotDatas;

    private Boolean drawPivotTable;

    private List<Integer> pivotTableBoundary;

    @Data
    public static class PivotColRow {
        private Integer index;
        private String name;
        private String fullname;
    }

    @Data
    public static class PivotValue {
        private Integer index;
        private String name;
        private String fullname;
        private String sumtype;
        private Integer nameindex;
    }
}
