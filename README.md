# WinCC 定时变量归档

本程序用VBS编写，在WinCC环境下将指定变量定时保存到指定的SQLite数据库中。

## 安装方法

以下需要在WinCC主机上安装32位和64位 SQLite ODBC Driver 程序。[下载页面](http://www.ch-werner.de/sqliteodbc/)

SQLite ODBC Driver 必须是32位，即使是64位操作系统。因为WinCC本身是32位应用程序。

### 在WinCC中安装

1. 将编译好的模块 tags_sqlite.bmo 复制到 WinCC 项目的 ScriptLib 文件夹下。
2. 按照需要的保存周期，复制对应编译好的动作文件 saveTags*.bac 至 WinCC 项目的 ScriptAct 文件夹下。见记录周期表格
3. 建立 D:\config\config.ini 配置文件，指定哪些变量以何种方式保存到数据库中。
4. WinCC 项目运行激活后，即可自动保存指定数值到数据库 D:\config\wincc.db 中。可用第三方工具读取数据库。

tags_sqlite.bmo 模块同时也提供 API(VBS)，可以在 WinCC 项目中根据需要取得指定时间和变量的历史值。

### 在WSH中安装

1. 将 sqlitags.wsh.js 和 所有的 *.cmd 文件复制到WinCC目标主机目录下，建议放在 D:\config\ 中。
2. 保证WinCC已激活
3. 配置好 D:\config\config.ini 文件，指定哪些变量以何种方式保存到数据库中。。（其它位置需要）
4. 双击 "start.cmd" 运行即可。

### 文件位置

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
MF33_M=GR_S7/Flow33.mass,GR_S7/Flow33.work_F,10minute
; 1#混合液体积累计
MF33_V=GR_S7/Flow33.volume,GR_S7/Flow33.work_F,10minute
; 1#纯油质量累积
MF33_O=GR_S7/Flow33.oil_mass,GR_S7/Flow33.work_F1,10minute
; 1#纯水质量累积
MF33_W=GR_S7/Flow33.water_mass,GR_S7/Flow33.work_F1,10minute
; 2#混合液质量累计
MF34_M=GR_S7/Flow34.mass,GR_S7/Flow34.work_F,1hour
; 2#混合液体积累计
MF34_V=GR_S7/Flow34.volume,GR_S7/Flow34.work_F,1hour
; 2#纯油质量累积
MF34_O=GR_S7/Flow34.oil_mass,GR_S7/Flow34.work_F1,1hour
; 2#纯水质量累积
MF34_W=GR_S7/Flow34.water_mass,GR_S7/Flow34.work_F1,1hour
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
