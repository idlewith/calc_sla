
# 安装Python

下面是Python的直接下载地址，记得勾选加入环境变量，其他都是点击下一步

https://repo.huaweicloud.com/python/3.9.7/python-3.9.7.exe


# 配置镜像地址

Windows路径为：C:\Users\<UserName>\pip\pip.ini

```
[global]
index-url = https://repo.huaweicloud.com/repository/pypi/simple
trusted-host = repo.huaweicloud.com
timeout = 120
```

# 安装依赖

```
pip install -r requirements.txt
```

# 用法

打开`template.xlsx`模板文件，在第一、二列输入事件的开始、结束时间即可

运行命令

```angular2html
python calc_sla.py
```
 
结果会输出到`result.xlsx`文件

# [可选] `config.json`配置文件

可以修改午休、上下班时间和输入输出的excel文件名

**午休时间**

项|含义
-|-
launch_break_start|午休开始时间
launch_break_end|午休结束时间
launch_break_delta_hour|午休需要减掉的时间

**上下班时间**

项|含义
-|-
off_duty_first_start|下班开始时间
off_duty_first_end|今天下班结束时间
off_duty_second_start|第二天下班开始时间
off_duty_second_end|第二天下班结束时间

**输入输出的excel文件名**

项|含义
-|-
input_excel_file_name|输入的excel文件名
input_excel_sheet_name|输入的excel文件需要读取的Sheet名
output_excel_file_name|输出的excel文件名


