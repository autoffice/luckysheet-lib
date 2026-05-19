# 已知问题

## Luckysheet导出为Excel存在的已知问题
* MATCH函数的参数默认值不一样，例如MATCH("dumpling",D21:D25)，在luckysheet中为MATCH("dumpling",D21:D25,0)，在wps中为MATCH("dumpling",D21:D25,1)
* WPS SUBTOTAL 函数无法叠加计算单元格内是函数表达式的情况
* Luckysheet迷你图自定义函数（LINESPLINES、COLUMNSPLINES、STACKBARSPLINES、PIESPLINES、BARSPLINES、AREASPLINES、TRISTATESPLINES、STACKCOLUMNSPLINES、DISCRETESPLINES）导出到Excel时无法解析，这些函数是Luckysheet前端专属，Excel不支持
* 不支持内嵌表格样式
* 图表导出为有损转换，仅保留基础结构
* 数据透视表导出为有损转换，仅保留基础结构
* 迷你图通过Excel扩展保存，有损

## Excel导入到Luckysheet存在的已知问题
* Excel中通过嵌入方式导入的图片不支持导入luckysheet，普通插入的图片支持
* POI的FormulaParser解析Luckysheet自定义函数时会输出WARN日志（Name 'XXX' is completely unknown），属于预期行为不影响功能