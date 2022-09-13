# WinCC 定时变量归档

本程序用VBS编写，在WinCC环境下将指定变量定时保存到指定的SQLite数据库中。

## 使用方法

1. 在WinCC主机上安装32位 SQLite ODBC Driver 程序，[下载页面](http://www.ch-werner.de/sqliteodbc/)
2. 将编译好的模块 tags_sqlite.bmo 复制到 WinCC 项目的 ScriptLib 文件夹下。
3. 按照需要的保存周期，复制对应编译好的动作文件 saveTags*.bac 至 WinCC 项目的 ScriptAct 文件夹下。
4. 建立 D:\config\config.ini 配置文件，指定哪些变量以何种方式保存到数据库中。
5. WinCC 项目运行后，自动保存指定数值到数据库 D:\config\wincc.db 中。可用第三方工具读取数据库。

tags_sqlite.bmo 模块同时也提供 API(VBS)，可以在 WinCC 项目中根据需要取得指定时间和变量的历史值。

注意：

* SQLite ODBC Driver 必须是32位，即使是64位操作系统。因为WinCC本身是32位应用程序。
* 如果配置文件不放置在 D:\config\config.ini，需要编辑 tags_sqlite 模块的 `CONF` 常量。
* 如果数据库文件不放置在 D:/config/wincc.db，需要编辑 tags_sqlite 模块的 `DBF` 常量。

## config.ini 语法

* 行首使用 ; 符号代表这一行为注释
* `[节名]` 代表一节，真到下一个节名前都是该节的配置。
* 本程序使用 `[tags]` 节指定WinCC要保存到sqlite数据库的变量
* `[tags]` 节下属每行是变量定义，定义一个WinCC变量如果保存到数据库中

### 变量定义格式

语法为：
`[归档变量名]=[WinCC变量],[有效性变量],[记录周期]`

* 归档变量名: 用在SQLite中的变量名称，SQLite单独使用变量名的原因：
  1. SQL中必须是标准变量标识符，WinCC变量名不合要求
  2. 归档变量能维持独立性，这样源WinCC变量名即使变化也不影响SQLite历史数据
* WinCC变量: WinCC变量管理中的名称
* 有效性变量: 也是 WinCC 变量，用来指示该变量值否有效。比如modbus的通讯正常指示变量。
* 记录周期: 可以填写的值
  * 1minute (不建议使用1minute，由于WinCC无法定时整分钟，会导致误差)
  * 10minute
  * 30minute
  * 1hour
  * 12hours
  * 1day
  * 1month

### 例子

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

可以在 WinCC 的画面或动作中使用该API，获得历史数值

```VBS
getHisTag(tagname, datastr)
```

* tagname 为[tags]节等号左侧的归档变量名
* datastr 为日期字符串

例：

```VBScript
Dim tag, currValue, prevValue
Set tag = HMTRuntime.Tags("半小时流量")
currValue = getHisTag("MF33_M", "2022-8-1 19:00")
prevValue = getHisTag("MF33_M", "2022-8-1 18:30")
tag.Write currValue - prevValue
```
