# EmailSpider


# Python3爬取outlook邮件汇总到Excel并将Excel发送到Outlook邮箱
## 功能：
通过爬取outlook邮件之后，根据发送的邮件标题模版，
汇总到Excel表格中
具体操作部分参考

## 需要修改的部分：
 1. **第一部分数据（代码第156行）**
        找到url对应的header中的cookie
 2. **第二部分数据（代码168和169行两个参数的值）**
      找到url2 对应header中的cookie 
      和X-OWA-CANARY值并替换，另外两个参数不变
## 发送邮件SendOutlook注意的部分：
 1. **安装使用python -m pip install pypiwin32，这个模块就包含win32com执行命令**
        python -m pip install pypiwin32
 2. **安装pypiwin32问题参考：https://zhangvalue.blog.csdn.net/article/details/103308248**
      
    