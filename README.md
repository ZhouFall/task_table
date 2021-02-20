# task_table
make work time data readable

1. 将从网上导出的excel表格，用python解析数据并绘制成图形
2. 用到的模块，画图模块， pandas


3参考网站：
https://zmister.com/archives/477.html
https://zmister.com/

4.Pycharm 安装QT环境总结：
	1. pip安装QT环境
	pip install pyqt5
	pip install pyqt5-tools
	2. 添加external Tools
	打开settings->Tools->External Tools点击“+”
	命名为QtDesigner,program为designer.exe(C:\Python37\Lib\site-packages\qt5_applications\Qt\bin)
	工作路径选当前文件夹
	3. 添加PyUIC
	路径为python.exe
	参数为-m PyQt5.uic.pyuic  $FileName$ -o $FileNameWithoutExtension$.py
	工作路径选当前文件夹
	4. 添加Pyrcc:
	路径为python37/Scripts/pyrcc5.exe
	参数为$FileName$ -o $FileNameWithoutExtension$_rc.py
	工作路径选当前文件夹
以上步骤相当于添加了快捷方式，在tools, external tools中可以调出来


2021/2/19
初步完成了数据整理和画图功能
Todo:
增加工时检查规则：比如总时长>8,却没有加班的，以及填写了加班，总共时为8的。
按照加班类型分类，无加班的，工时<=8,延时加班且无请假的，总共时=8+加班时长（如果有请假还需要减去）， 休息日加班的，总时长=加班时长
增加费用统计处理，绘制成图形
      
      
Todo:
将output中的内容打包通过邮件自动发送出去，可以单独开发

Todo:
研究定时功能

Todo:
爬虫从网站上爬取数据，并下载，优先级最低
