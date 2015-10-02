简介：
用python写了个脚本来爬数据，可以自动补全产品名字，包装信息和电子书信息，图片检查和搜索还需要自己检查，
不过可以生成一个html文件，方便搜索检查，支持linux，windows可能会有点问题

事前准备：
windows环境下：
1.安装Pyhton 2.7并配置相关环境变量，http://jingyan.baidu.com/article/7908e85c78c743af491ad261.html
2.安装PIP，http://blog.chinaunix.net/uid-12014716-id-3859827.html
3.下载并安装pywin32,
	64位下这个：http://sourceforge.net/projects/pywin32/files/pywin32/Build%20219/pywin32-219.win-amd64-py2.7.exe/download
	32位下这个：http://sourceforge.net/projects/pywin32/files/pywin32/Build%20219/pywin32-219.win32-py2.7.exe/download
4.安装PIL（如果装不上这个就转linux平台吧）,http://effbot.org/downloads/PIL-1.1.7.win32-py2.7.exe
3.以管理员权限运行命令行窗口,输入：
	pip install scrapy
	pip install xlrd
	pip install xlwt
	pip install xlutils
5.安装git

开始步骤
1.在git bash下输入git clone https://github.com/codefarmer-cyk/vonaweb_scrapy.git
2.把你的作业文件（*.xls）复制到file文件夹里，如：2200-2399逸逵.xls
3.修改vonaweb/spider目录下的VonaSpeder.py的22行，把2200-2399逸逵.xls替换成你的文件名
4.修改result.py的42行的2200-2399逸逵.xls替换成你的文件名
5.在第一个vonaweb目录下打开命令行窗口，输入scrapy crawl vona -o vona.json
6.等爬虫脚本跑完后，输入python result.py
7.在file文件夹里会生成item.html和result.xls文件，item.html里可以看到每个商品的名字，图片和对应查询的超链接，result.xls已经
	把产品名字，包装信息那几列补全
	
注意事项：
1.生成的xls文件格式会有些不同，要修改：
	列的宽度，
	单元格->自动换行
	数据—>有效性->列表，然后把那下拉菜单的值填上去
2.result.xls有些列的值为空或图片不存在可能是连接超时导致的，可以自己再检查一下
2.有bug可向我反馈	