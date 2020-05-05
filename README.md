# Pandas&Excel常用操作
 使用pandas数据分析Excel常用的一些方法
<h1>操作说明</h1>
<h3>1.toXlsx(path)：</h3>
	<p>将带有xlm的xls文件转换为pandas可以直接读取的xlsx文件。
	<p>path：若为'.'，则转换当前目录下；如果为文件夹绝对路径则转换绝对路径下的xls
<h3>2.heBingExcel(path, finalName, headCounts, tailCounts)</h3>
<p>将path路径下的所有excel文件(后缀为xls或xlsx的文件)合并成只有一个sheet的excel文件
<p>path为绝对路径
<p>headCounts为表头有几行
<p>tailCounts为表尾有几行
<h3>3.统计某些列总计
<h3>4.查找特定列包不包含多字符，使用或不使用lambda表达式
<h3>5.多条件筛选
<h3>6.多表联合查询
<h3>7.可视化