package io.github.autoffice.luckysheet.model.sheet;

import lombok.Data;

/**
 * 表格数据验证配置实体
 * 对应字段规则：严格按照你提供的 dataVerification 结构生成
 */
@Data
public class DataVerification {

    /**
     * 类型
     * dropdown(下拉列表)
     * checkbox(复选框)
     * number(数字)
     * number_integer(数字-整数)
     * number_decimal(数字-小数)
     * text_content(文本-内容)
     * text_length(文本-长度)
     * date(日期)
     * validity(有效性)
     */
    private String type;

    /**
     * 条件类型（根据 type 不同而不同）
     */
    private String type2;

    /**
     * 条件值1
     * 下拉框：选区/逗号分隔字符串
     * 有效性：可为空
     * 其他：数值/字符串
     */
    private Object value1;

    /**
     * 条件值2
     * 仅 checkbox / bw / nb 时需要
     * 日期/数值必须 >= value1
     */
    private Object value2;

    /**
     * 自动远程获取选项 默认 false
     */
    private Boolean remote = false;

    /**
     * 输入无效时禁止输入 默认 false
     */
    private Boolean prohibitInput = false;

    /**
     * 选中单元格时显示提示语 默认 false
     */
    private Boolean hintShow = false;

    /**
     * 提示语文本（hintShow=true 时必须配置）
     */
    private String hintText;

    /**
     * 是否勾选（仅 type=checkbox 时需要）
     */
    private Boolean checked;

    /**
     * 默认空对象
     */
    public DataVerification() {
    }
}
