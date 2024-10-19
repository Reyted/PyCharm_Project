from pathlib import Path

while True:
    floder=input('请输入要搜索得文件路径：')
    floder=Path(floder.strip())
    if floder.exists():
        break
    else:
        print('您输入得文件路径有误！')

search=input('请输入搜索内容：')
list=list(floder.rglob(f'*{search}*'))
if not list:
    print('未找到相关内容')
else:
    for i in list:
        print(i)
