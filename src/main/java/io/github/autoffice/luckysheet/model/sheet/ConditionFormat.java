package io.github.autoffice.luckysheet.model.sheet;

import lombok.Data;

import java.util.List;


/**
 * luckysheet 条件格式配置信息
 *
 * @see <a
 * href="https://dream-num.github.io/LuckysheetDocs/zh/guide/sheet.html#luckysheet-conditionformat-save">
 * luckysheet-conditionformat-save</a>
 */
@Data
public class ConditionFormat {
    /**
     * 突出显示单元格规则和项目选区规则
     */
    private ConditionFormatType type;
    /**
     * 条件应用范围
     */
    private List<Range> cellrange;
    /**
     * 格式,根据不同的{@link ConditionFormat#type},实际对象类型不同
     */
    private Object format;
    /**
     * 类型
     *
     * <p>Detailed settings,comparison parameters
     */
    private String conditionName;
    /**
     * 条件值所在单元格
     *
     * <p>Detailed settings,comparison range
     */
    private List<Range> conditionRange;
    /**
     * 自定义传入的条件值
     */
    private List<Object> conditionValue;
}
