# CATIA 自动化BOM程序

本程序通过EXCEL中VBA和控件实现CATIA产品设计中：

- 通过初始化，实现零件的属性自定义（iMass, iDensity, iTthickness and iMaterial），同时也为所有级别的总成和零件创建了生成BOM时需要的参数（mass），以及这些参数在零件内的自动化计算或生成；

- 通过读取零件属性和参数到excel，或者读取excel中的值，实现对零件属性和参数的读取\修改\批量修改;

- 通过递归实现对所有总成\分总成的重量计算；

- 通过递归实现将产品的所有子零件\部件的信息批量导出至Excel，实现BOM的自动化；

# 使用注意事项

- CATIA必须具有knowledge advisor模块lisence，并安装有VBA扩展；
- excel请调用CATIA的dll库，并启用信任和宏的权限；
