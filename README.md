# catword


护网开始了，钓鱼文件满天飞，万一是 0day 怎么办？ 

catword: 终端查看 doc/docx 文件，纯 Python 实现，支持 Windows/Linux/MacOS 系统。



```
# 目前唯一的依赖是 olefile
pip install olefile

python catword.py xx.doc[x]


# linux style

python catword.py xx.doc[x] | less

```





