1	总体说明:	
      通过excel VBA获取NetSuite的数据用于分析。

2	实现方式:
      通过VBA调用获取NetSuite数据后以json方式存储到本地，让后通过powerquery把json文件解析到excel表中
      
3	使用:
      1、使用前请填写相关参数；然后在查询清单中，根据清单的样例，填写需要查询的内容
      2、将光标放在想要执行的查询名称上，点击执行。
      3、任何时候不会更新NetSuite的数据
      4、执行过程中，除了有弹出框，执行过程的错误消息会在ErrorMsg的页签中出现
      5、如果你想修改查询结果的字段名称，请在Query字段描述的地方修改"
      
4	参数说明:
        1、参数中KeyFielName是指定的密钥文件，根据NetSuiteKey.json文件填写NetSuite的登录信息
        2、参数中Json文件目录是存放从NetSuite中获取数据后以json文件存储的目录

5	问题反馈:
    ziphei@outlook.com
