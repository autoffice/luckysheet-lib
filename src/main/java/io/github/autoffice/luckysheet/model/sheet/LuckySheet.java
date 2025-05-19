package io.github.autoffice.luckysheet.model.sheet;

import io.github.autoffice.luckysheet.model.cell.CellData;
import io.github.autoffice.luckysheet.model.image.SheetImage;
import lombok.Data;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Data
public class LuckySheet {

    /**
     * new 对象,并对关键值初始化
     */
    public LuckySheet() {
        this.config = new SheetConfig();
        this.celldata = new ArrayList<>();
        this.images = new HashMap<>();
    }

    /**
     * 工作表名称
     */
    private String name;

    /**
     * 工作表颜色,工作表名称下方会有一条底部边框。hex值, eg: #f20e0e
     */
    private String color;

    /**
     * 表格行高、列宽、合并单元格、边框、隐藏行等设置
     * <p>
     * 如果为空,只能为空对象,不能为null
     */
    private SheetConfig config;

    /**
     * 工作表索引，作为唯一key值使用，新增工作表时会自动赋值一个随机字符串。注意index不是工作表顺序，和order区分开
     */
    private String index;

    /**
     * 此sheet页面的缩放比例, 为0~1之间的二位小数数字. eg: 1、0.1、0.56.  default 1
     * <p>
     */
    private Double zoomRatio;

    /**
     * 图表
     */
    private List<SheetChart> chart;

    /**
     * 工作表的下标，代表工作表在底部sheet栏展示的顺序，新增工作表时会递增，从0开始
     */
    private Integer order;

    /**
     * 是否隐藏，0为不隐藏，1为隐藏
     */
    private BoolStatus hide;

    /**
     * 列数,包括空列
     */
    private Integer column;

    /**
     * 行数,包括空行
     */
    private Integer row;

    /**
     * 激活状态，仅有一个激活状态的工作表， 1,选中; 0,未选中
     */
    private BoolStatus status;

    /**
     * 初始化使用的单元格数据
     */
    private List<CellData> celldata;

    /**
     * 选中的区域(支持多选)
     */
    private List<Range> luckysheet_select_save;

    /**
     * 图片
     */
    private Map<String, SheetImage> images;

    /**
     * 公式链 TODO 待支持 公式链是一个由用户指定顺序排列的公式信息数组，Luckysheet会根据此顺序来决定公式执行的顺序。
     *
     * @see <a
     * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#calcchain">calcChain</a>
     */
    private List<CalcChain> calcChain;

    /**
     * 左右滚动条位置
     * <p>
     * ECMA-376中无对应属性
     */
    private Integer scrollLeft;

    /**
     * 上下滚动条位置
     * <p>
     * ECMA-376中无对应属性
     */
    private Integer scrollTop;

    /**
     * 是否数据透视表
     */
    private Boolean isPivotTable;
    /**
     * 数据透视表
     */
    private SheetPivotTable pivotTable;

    /**
     * 条件格式配置信息，包含多个条件格式配置对象的一维数组，
     */
    private List<ConditionFormat> luckysheet_conditionformat_save;

    /**
     * 筛选范围。一个选区，一个sheet只有一个筛选范围，类似luckysheet_select_save。
     * 如果仅仅只是创建一个选区打开筛选功能，则配置这个范围即可，
     * 如果还需要进一步设置详细的筛选条件，则需要另外配置同级的 filter 属性。
     */
    private Range filter_select;

    /**
     * 筛选的具体设置，跟filter_select筛选范围是互相搭配的。当你在第一个sheet页创建了
     * 一个筛选区域，通过luckysheet.getLuckysheetfile()[0].filter也可以看到第
     * 一个sheet的筛选配置信息。
     */
    private Object filter;

    /**
     * 替颜色配置
     */
    private List<AlternateFormat> luckysheet_alternateformat_save;

    /**
     * 默认行高(pixel)
     */
    private Short defaultRowHeight;

    /**
     * 默认列宽(pixel)
     */
    private Short defaultColWidth;

    /**
     * 自定义交替颜色，包含多个自定义交替颜色的配置
     */
    private List<AlternateFormat.AlternateFormatValue> luckysheet_alternateformat_save_modelCustom;

    /**
     * 是否显示网格线. 1, 显示; 0,隐藏.
     */
    private BoolStatus showGridLines;
}
