# 基于Python的实验室自动化测试系统

如果能用最简单的方法（Python）实现需求，就无需使用更复杂的方法。杀鸡焉用牛刀（Visual Studio）。

## 组网图

![connect](https://github.com/yugiyx/python_happy_test/blob/master/template%20and%20diagram/device%20connect%20diagram.png)


## 特点

* 整套系统绿色，无需安装，硬盘占用空间小。配置要求低。
* 黑盒化。经过简单训练后，任何人都可以操作。
* 现实设备和仪器虚拟化。现实怎么操作设备仪器，软件就怎么操作。理论可以无限扩展。
* 使用excel作为参数输入和结果输出。更容易测试用例结合。

## 必备组件

### Python及相关库
* [python3.4.3](https://www.python.org/downloads/release/python-343/)
* [pyvisa 1.8](https://pypi.python.org/pypi/PyVISA/1.8)
* [pyserial 3.0](https://pypi.python.org/pypi/pyserial/3.0)
* [xlrd 1.1.0](https://pypi.python.org/pypi/xlrd/1.1.0)
* [xlwt 1.3.0](https://pypi.python.org/pypi/xlwt/1.3.0)
* [xlutils 2.0.0](https://pypi.python.org/pypi/xlutils/2.0.0)

### VISA
* [National Instruments’s VISA](http://www.ni.com/visa/)

### 安装说明：
考虑很多实验室仍然使用Windows XP系统，并且没有网络或者没有管理员权限，不能安装任何程序（作者公司提供的就是这种工作环境，吐槽一下:weary:）。
* Python选择的是最后一个支持xp版本3.4.3。pyserial选择的3.0。其他库没有特别要求，使用最新版即可。
* 可以在能够安装python和库的机器上，安装好全部库，打包整个python文件夹，直接拷贝到其他机器使用，在Sublime等文本编辑器中增加路径即可绿色化。


## 程序组件说明


* `edfa.py`	待测设备命令接口
* `instrument.py`	实验室仪器虚拟化接口
* `pd_test.py`	PD测试程序
* `nf_test.py`	NF测试程序
* `test_case.xls`	全部测试的测试条件（输入）
* `test_result_pd.xls`	PD测试结果（输出）
* `data_calc.py`	NF测试计算加汇总程序
* `test_result_all.xl`	NF测试结果汇总（输出）


## 配置程序说明
以PD测试的为例。以下代码需要根据实际情况配置。
```python
# Select Test PD ID
test_pd = 'APC1'
test_edfa = 'E1'

# Config test enviorment
tls_1510_pd = ['PD2']
tls_1610_pd = ['PD16']
out_pd = ['PD9', 'PDT1', 'PDT2', 'APC1', 'APC2']
edfa = edfa.EdfaOFP2(2)
att_pm = instrument.AttPmN7752A(30)
pm = instrument.Pm8163A(21)
switch = instrument.SwZHDIY(6)
file = instrument.DataProcess()

# Initial test module.Modified according to different modules
if test_pd in tls_1510_pd:
    pm.set_pm_wavlength(1510)
    att_pm.set_att_wavlength(1510)
    switch.set_sw('TLS_EDFA_PM')
    pin_offset = 0
    pout_offset = 1.8
elif test_pd in tls_1610_pd:
    pm.set_pm_wavlength(1610)
    att_pm.set_att_wavlength(1610)
    switch.set_sw('TLS_EDFA_PM')
    pin_offset = 0
    pout_offset = 1.8
else:
    pm.set_pm_wavlength(1550)
    att_pm.set_att_wavlength(1550)
    switch.set_sw('MLS_EDFA_PM')
    pin_offset = 1
    pout_offset = -1
att_pm.enable_apc_mode(1)
att_pm.set_apc_value(-10)
edfa.set_edfa_mode(test_edfa, 'AGC')
edfa.write_reg('41', '11')  # Modified
edfa.set_edfa_gain(test_edfa, 18)

# Read set values from excel file
all_set_values = file.open_config('test_case.xls', test_pd)
pd_report = [[], [], []]
pm_report = []
print(all_set_values[0])
```

## 可借鉴方法

这里已经包括了特殊的产品，以及标准光学实验室仪器（频谱分析仪，激光器，衰减器，攻击率，电子开关）。理论能用串口、并口、网口、GPIB，USB控制的设备都能够参考此套代码。

## 下一步计划

* Fix：目前程序内部仍然有一处需要根据不同测试模块手工修改的缺陷。
* Fix：xlrd写入公式，必须手工打开excel并保存一次，才能被xlwt读取到计算值的缺陷。
* CR：增加示波器的代码。
* CR：找到合适的 1XN swtich，可以把被测端口全部连上，一次全部测完，无需手工拔插端口和修改代码选择被测端口。
* CR：将所有配置修改部分，全部放入excel表格，做到程序完全黑盒化。
* CR：优化代码，使程序更简洁优雅。

## 版本记录

[CHANGELOG](https://github.com/yugiyx/happy-send/blob/master/CHANGELOG.md)




