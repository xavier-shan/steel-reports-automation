# 钢贸报表自动化脚本

这是一套按你当前工作流整理好的自动化脚本，分为两部分：

- `scripts/price_daily.py`：市场价格日报
- `scripts/manager_weekly.py`：经理周报
- `scripts/export_sheet_snapshot.py`：可选，把 Excel 结果导成整页 PNG 预览图

## 这套脚本已经内置的业务规则

### 市场价格日报
1. 以模板/上一版日报为基础，保留原格式。
2. 输入当天价格后，优先更新显式输入的钢厂和钢种。
3. 新兴铸管、凌钢、徐钢/六安/大东海，会按上一版里和三钢的固定差价联动。
4. 所有 `40Cr` 会按各自和 `45#` 的固定差价自动联动。
5. 所有 `195` 线材会按各自和 `08Al` 的固定差价自动联动。
6. `涨跌` 按**当天相对上一版**重算：涨X / 跌X / 平。
7. `和元立差价` 按“泉州地区送到价”与元立对应品种送到成本逐项比较，自动写成：
   - `比元立送到成本低X`
   - `比元立送到成本高X`
   - `和元立送到成本持平`
8. 备注里的 `元立建议出厂价` 按 `三钢45# - 150` 生成。
9. 若模板里还没有 `徐钢，六安，大东海` 这行，会自动插在 `凌钢` 的下一行，写入：
   - 钢种：`45#`
   - 贸易商报价：`3580`
   - 泉州地区送到价：`3600`

### 经理周报
1. 固定逻辑：
   - `C5 <- F5`
   - `C6 <- F6`
   - `C7 <- F7`
2. 然后继续轮换：
   - `F5` 从圆钢客户列表取下一个
   - `F6`、`F7` 从线材客户列表依次取下两个
3. `福建广吉` 可固定在配置指定位置，不参与轮换。
4. `B` 列送到价按输入覆盖。
5. 轮换索引会写入状态文件，下一次继续接着转。

## 目录结构

```text
steel_reports_automation/
├── README.md
├── .gitignore
├── requirements.txt
├── run_examples.sh
├── config/
│   ├── manager_rotation_state.example.json
│   ├── manager_weekly_run.example.json
│   └── market_price_run.example.json
├── launchd/
│   ├── com.xavier.manager-weekly.plist
│   └── com.xavier.market-price-daily.plist
└── scripts/
    ├── export_sheet_snapshot.py
    ├── manager_weekly.py
    ├── price_daily.py
    └── utils.py
```

## 安装

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 用法

### 1）市场价格日报
先复制示例配置：

```bash
cp config/market_price_run.example.json config/market_price_run.json
```

然后把里面这些路径改成你的真实路径：
- `template_path`
- `runtime.snapshot_path`
- `runtime.output_dir`

再填当天价格：

```json
{
  "inputs": {
    "report_date": "2026-04-22",
    "changed_prices": {
      "三钢": {
        "45#": 3660
      },
      "元立": {
        "45#": 3520,
        "08Al": 3440,
        "195": 3340
      }
    }
  }
}
```

执行：

```bash
python3 scripts/price_daily.py --input config/market_price_run.json
```

### 2）经理周报
先复制配置和轮换状态：

```bash
cp config/manager_weekly_run.example.json config/manager_weekly_run.json
mkdir -p state
cp config/manager_rotation_state.example.json state/manager_rotation_state.json
```

然后修改：
- `template_path`
- `runtime.rotation_state_path`
- `runtime.output_dir`

填入这周送到价：

```json
{
  "inputs": {
    "report_date": "2026-04-24",
    "delivered_prices": {
      "B5": 3660,
      "B6": 3550
    }
  }
}
```

执行：

```bash
python3 scripts/manager_weekly.py --input config/manager_weekly_run.json
```

### 3）导出整张表 PNG 预览
这一步是可选的，适合你要求“整张表完整入图”的场景。需要先装：
- LibreOffice
- ImageMagick

执行：

```bash
python3 scripts/export_sheet_snapshot.py --xlsx ./output/市场价格日报4月22日.xlsx --png ./output/市场价格日报4月22日.png
```

## macOS 定时执行（可选）
`launchd/` 目录已经给了两个 plist 示例：
- 工作日 `09:40` 价格日报
- 周五 `09:45` 经理周报

你只需要把里面的绝对路径替换掉，然后执行：

```bash
launchctl unload ~/Library/LaunchAgents/com.xavier.market-price-daily.plist 2>/dev/null || true
cp launchd/com.xavier.market-price-daily.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.xavier.market-price-daily.plist

launchctl unload ~/Library/LaunchAgents/com.xavier.manager-weekly.plist 2>/dev/null || true
cp launchd/com.xavier.manager-weekly.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.xavier.manager-weekly.plist
```

## 目前这套脚本的边界
因为我现在**没有拿到你的真实模板文件**，所以这套代码是按“保留模板格式 + 自动识别表头 + 配置化运行”的方式写的。

也就是说：
- 逻辑已经按你现在的固定规则写进去
- 只要模板表头跟示例接近，就可以直接跑
- 如果你的真实模板里列名、工作表名、单元格位置有差异，我拿到模板后可以再帮你一次性贴合到位
