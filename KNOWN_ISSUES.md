# 已知问题

## Luckysheet导出为Excel存在的已知问题
* MATCH函数的参数默认值不一样 ，例如MATCH("dumpling",D21:D25)，再luckysheet中为MATCH("dumpling",D21:D25,0)，再wps中为MATCH("dumpling",D21:D25,1)
* WPS SUBTOTAL 函数无法叠加计算单元格内是函数表达式的情况
* luckysheet中的函数LINESPLINES(B3:B5,'pink',4,'avg','yellow','red','green',3)导出到excel中有兼容问题，报函数无法解析
* luckysheet中的函数COLUMNSPLINES(B3:B5,35,'red','green','auto','brown')导出到excel中有兼容问题，报函数无法解析
* 不支持条件格式
* 不支持内嵌表格样式
* 不支持生成线图
* 不支持数据透视表
* 不支持生成图表
* 不支持数据验证

## Excel导入到Luckysheet存在的已知问题
* Excel中通过嵌入方式导入的图片不支持导入luckysheet, 普通插入的图片支持