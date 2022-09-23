# WinCC 定时变量归档

本程序用VBS编写，在WinCC环境下将指定变量定时保存到指定的SQLite数据库中。

## 安装方法

1. 在WinCC主机上安装32位和64位 SQLite ODBC Driver 程序。[下载页面](http://www.ch-werner.de/sqliteodbc/)(SQLite ODBC Driver 必须是32位，即使是64位操作系统。因为WinCC本身是32位应用程序。)
1. 将本项目下 sqlitags.wsh.js、run.vbs 和 所有的 *.cmd 文件复制到WinCC主机的 D:\config\ 中。
1. 建立并配置 D:\config\config.ini 文件，指定哪些变量以何种方式保存到数据库中。
1. 在WinCC中配置启动脚本（也可以在激活WinCC后手动启动脚本）
1. 激活WinCC

### 在WinCC中配置启动脚本

在 WinCC 项目中启动  sqlitags.wsh.js，可以用以下二种办法之一：

1. 在定时动作中加入以下代码：

   ```VBScript
   Dim ws : Set ws = CreateObject("WScript.Shell")
   ws.CurrentDirectory = "D:\config"
   ws.Run "C:\WINDOWS\SysWOW64\CScript.exe D:\config\run.vbs start /b", 0
   ```

1. 在WinCC项目启动中加入`C:\WINDOWS\SysWOW64\CScript.exe D:\config\run.vbs restart /b`

### 在WSH中手动启动

激活WinCC后运行 `./start.cmd /b`。

### 说明

随时可以用 `./list.cmd` 查看是否变量归档在运行。

正常情况下按默认位置放置配置文件和生成数据库文件。如果因没有D盘等原因要修改文件位置，可以：

* WinCC 版本下，编辑 tags_sqlite 模块的 `CONF` 常量与 `DBF` 常量。
* WSH 运行环境下修改 sqlitags.wsh.js 中的 `CONF` 变量 `DBF` 变量量。

## config.ini 语法

* 行首使用 ; 符号代表这一行为注释
* `[节名]` 代表一节，真到下一个节名前都是该节的配置。
* `[tags]` 节定义有哪些 WinCC 变量保存到sqlite数据库中
  * 本节所有变量不能有计算式，一对一保存数值历史
  * 下属每行是变量定义，定义具体某个WinCC变量如何保存

### [tags]节变量定义格式

语法为：

`[归档变量名]=[WinCC变量],[有效性变量],[记录周期]`

归档变量名: 用在SQLite中的变量名称，SQLite单独使用变量名的原因：

1. SQL中必须是标准变量标识符，WinCC变量名不合要求
2. 归档变量能维持独立性，这样源WinCC变量名即使变化也不影响SQLite历史数据

WinCC变量: WinCC变量管理中的名称

有效性变量: 也是 WinCC 变量，用来指示该变量值否有效。比如modbus的通讯正常指示变量。

记录周期: 可设为以下表格中值。WSH版无需设置动作，自己有定时器。

| 周期值       | WinCC版需要设置的动作 | 描述                                  |
| --------- | ------------- | ----------------------------------- |
| 1minute   | saveTagsPer1N | 1分钟 (WinCC版不建议使用1minute，无法整分钟定时)    |
| 10minutes | saveTagsOn10N | 00:00 10:00 20:00 30:00 40:00 50:00 |
| 30minutes | saveTagsOn10N | 半小时                                 |
| 1hour     | saveTagsOn1H  | 整点                                  |
| 2hoursO   | saveTagsOn1H  | 奇数整点                                |
| 2hoursE   | saveTagsOn1H  | 偶数整点                                |
| 12hours   | saveTagsOn12H | 每日 00:00 与 12:00                    |
| 1day      | saveTagsOn12H | 每日 00:00                            |
| 1month    | saveTagsOn1M  | 每月1日 00:00                          |

注：同一个WinCC变量，如果多个时间周期存档需求，只需定义最小周期那一个即可

### [tags]节例子

```ini
[tags]

; ## 稠油质量流量计工程用
; 1#混合液质量累计
MF33_M=GR_S7/Flow33.mass,GR_S7/Flow33.work_F,10minutes
; 1#混合液体积累计
MF33_V=GR_S7/Flow33.volume,GR_S7/Flow33.work_F,10minutes
; 1#纯油质量累积
MF33_O=GR_S7/Flow33.oil_mass,GR_S7/Flow33.work_F1,10minutes
; 1#纯水质量累积
MF33_W=GR_S7/Flow33.water_mass,GR_S7/Flow33.work_F1,10minutes
; 2#混合液质量累计
MF34_M=GR_S7/Flow34.mass,GR_S7/Flow34.work_F,1hour
; 2#混合液体积累计
MF34_V=GR_S7/Flow34.volume,GR_S7/Flow34.work_F,1hour
; 2#纯油质量累积
MF34_O=GR_S7/Flow34.oil_mass,GR_S7/Flow34.work_F1,1hour
; 2#纯水质量累积
MF34_W=GR_S7/Flow34.water_mass,GR_S7/Flow34.work_F1,1hour
```

### [OS] 节变量定义格式

[OS] 节为运算变量，运算后的结果保存至WinCC变量中。

只有WSH版的 SQLiTags 脚本支持 [OS] 节，同时只支持数值类型变量。

与tags节相反，本节定义了需要计算表达式后存入WinCC变量，表达式中可能有SQLite读取的归档历史数值

格式：`WinCC变量=表达式,生成时间间隔`

表达式中：

* 变量用{}括起来，即可以是归档变量，也可以是WinCC当前变量
* {Y} {M} {D} {H} {N} {W} 代表WinCC系统的当前年月日时分周等数值
* {WinCC_tag_name} 非上述日期变量则代表WinCC变量值
* {DB_tag_name,年,月,日,时,分,秒} 带 `,` 分隔的变量，为归档变量，含日期部分
  * 日期的各部分，支持对应日期变量 Y M D H N S 开头的加减运算
  * 天是以1开始，0代表上个月的最后一天
  * 其它日期部分用-1代表上个周期的最后时刻，比如-1分代表上一小时的59分。
  * 例：{tname,Y,M,0,23,59,0}表示上个月最后一天的23:59
* 除归档变量的日期部分外，`{}`中不得有运算符

生成时间间隔可用的值为：

* 1minute
* 10minute
* 30minute
* 1hour
* 2hoursO 奇数整点
* 2hoursE 偶数整点
* 12hours
* 1day
* 1month

### [OS]节例子

```ini
[OS]
;混合液质量差（混合液质量累计值的差值，单位t）
Flow33_mass_diff_1N={MF33_M,Y,M,D,H,N-N%10,0}-{MF33_M,Y,M,D,H,N-1-N%10,0},10minutes
Flow33_mass_diff_2H={MF33_M,Y,M,D,H,0,0}-{MF33_M,Y,M,D,H-2,0,0},2hoursE
Flow33_mass_diff_1D={MF33_M,Y,M,D,0,0,0}-{MF33_M,Y,M,D-1,0,0,0},1day
```

## API

如果安装在 WinCC 中，可以在 WinCC 的画面或动作中使用 API 获得历史数值。

```VBS
getHisTag(varname, datastr)
```

* varname 为 SQLite 中的归档变量名
* datastr 为日期字符串

例：

```VBScript
Dim tag, currValue, prevValue
Set tag = HMTRuntime.Tags("半小时流量")
currValue = getHisTag("MF33_M", "2022-8-1 19:00")
prevValue = getHisTag("MF33_M", "2022-8-1 18:30")
tag.Write currValue - prevValue
```
