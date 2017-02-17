## 文件说明
### beinongda.xls	
- 搜狗地图导出的文件

### model.xls		
- 需要提交的文件，作为模版用

### sheetlist.txt	
- 从model.xls中提取sheet，组成列表。只对这个文件中的名字数据处理，其他不处理。
- 由sheetlist.py生成
- 如有人员增减，直接编辑这个文件即可

### location.txt	
- 从beinongda.xls导出的位置信息
- 由Personal.py生成

### no_signal.txt	
- 从beinongda.xls导出的位置异常数据（导出的文件是没有相关日期，或相关时间段显示“没有收到卫星信号”）
- 由Personal.py生成
		注：这样的数据
		张华伟 2014-05-05 08:29:10 没有收到卫星信号 2014-05-05 18:27:50	没有收到卫星信号	全天无有效数据
		表示：当天全天都显示“没有收到卫星信号”，按“全天无数据”处理

### no_signal2.txt	
- 对no_signal.txt的处理， 记录：超过30分钟及文件上没有的时间
- 由nosignal.py生成

### no_signal3.txt	
- 对no_signal.txt的处理， 记录：显示“没有收到卫星信号”时间段不足30分钟
- 由nosignal.py生成

### report.xls		
- 生成的报表
- 由report.py生成

### report_jicheng.xls	
- 生成的统计报表
- 由report_jicheng.py生成

### namelist.txt 
- 按照模版顺序的人名

### sheetlist.txt
- Excel中sheet的名称

## 使用说明
1. 运行sheetlist.py，提取Excel中的sheet，保存到sheetlist.txt中
2. 运行Personal.py，生成位置文件location.txt和无信号文件
3. 运行nosignal.py，生成无信号超过最大时间的no_signal2.txt和短于最大时间的no_signal3.txt，方便及时通知本人
4. 运行report.py，生成报表report.xls文件
5. 运行report_jicheng.py，生成统计报表 
