# 13_结算整理

两个脚本：建空 JSON、汇总成 Excel。内容自己填。

## 1. 创建 JSON（已有文件不覆盖已填内容）

```bash
cd /Users/wyattsun/Desktop/code/13_结算整理
python 00_创建JSON.py
python 00_创建JSON.py --dry-run
```

在每个 `THCC-xxxx-xxx_项目名` 文件夹下生成 `THCC-xxxx-xxx.json`。  
只自动写入项目编码、项目名称；缺字段时补空键，不改已填写内容。

## 2. 汇总 Excel

```bash
python 01_汇总输出Excel.py
```

默认输出：

`/Volumes/PC-SunYongtao/工作临时文件存放包/运行考核部/业主结算额/业主结算汇总.xlsx`

## 汇率

JSON 里有「币种」「汇率」。汇总时计算「有效汇率」和「金额_泰铢」。

| 情况 | 有效汇率 |
|------|----------|
| 「汇率」填了数字 | 用填写值 |
| 不填 / 空，币种泰铢或空 | 1 |
| 不填 / 空，币种 CNY | 5 |
| 不填 / 空，币种 USD | 32 |

折合泰铢 = 原币金额 × 有效汇率。
