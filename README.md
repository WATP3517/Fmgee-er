# Fmgee-艾尔登法环fmg翻译工具-xml转换-excel编辑

·将艾尔登法环中xml格式的fmg文件转换为表格，进行翻译的转换工具

·XML转换为EXCEL表格
	
 	将.xml文件拖至Fmgee_er.exe上方，可将文本内容转换为.xlsx格式的excel表格
	表格中自动折叠%null%空文本行，未翻译的纯英文文本行会用红色背景标出

·EXCEL表格转换回XML

  
	将.xlsx文件拖至Fmgee_er.exe上方，可将excel表格转换回原格式的xml文件

·批量操作
  
	选择多项文件拖至Fmgee_er.exe上方，可对文件进行批量转换

·使用教程

	1.下载解包工具yabber https://github.com/JKAnderson/Yabber
	
	2.将fmg/zhocn文件夹下的dcx文件拖至yabber.exe上方解包
      得到两个dcx后缀的文件夹
		
	3.打开文件夹\GR\data\INTERROOT_win64\msg\Zhocn
      将其中的fmg文件拖至yabber.exe上方解包
      得到解包出的xml格式文件
  
	4.将xml文件拖至Fmgee_er.exe上方转换为excel表格
  
	5.打开表格进行翻译
  
	6.将翻译完成的表格拖至Fmgee_er.exe上方转换回xml文件
  
	7.将xml文件拖至yabber.exe上方打包回fmg文件
  
	8.退回fmg/zhocn文件夹，将dcx文件夹拖至yabber.exe上方打包回dcx文件
