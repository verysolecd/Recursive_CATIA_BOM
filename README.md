本程序通过EXCEL中VBA和控件实现CATIA产品设计中：
- 通过初始化，实现零件的属性自定义（mass, density, tihickness and materail），同时也为总成、分总成写入了进入BOM的参数（mass）
- 通过读取零件属性和参数到excel，或者读取excel中的值，实现对零件属性和参数的读取\修改\批量修改;
- 通过递归实现对所有总成\分总成的重量计算；
- 通过递归实现将产品的所有子零件\部件的信息批量导出至Excel，实现BOM的自动化；
