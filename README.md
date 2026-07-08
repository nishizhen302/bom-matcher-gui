# BOM匹配助手

BOM 清单与物料编码表（Y06/Y12 料号体系）自动匹配工具：桌面 GUI（pywebview）+ 可独立调用的匹配引擎，
输出带料号、聚合汇总、四色标注的 BOM Excel，未匹配项支持界面内人工复核后一键重导出。

## 功能

- **智能匹配**：电阻/电容/电感/磁珠/连接器/IC/二极管/保险丝/晶振等多类器件；
  先把 BOM 写法与编码表规格解析成结构化字段（值/电压/精度/封装/尺寸/型号），再做硬约束过滤+软打分排序
- **物料收敛**：同容值+封装的电容全局统一选料（高压兼容替代低压），减少物料种类
- **输出可读**：料号排序；重复组灰色+绿色汇总行；带颜色的行集中在表尾——
  橙色[需确认] → 杏色非采购项（安装孔/测试点）→ 黄色未匹配 → 蓝色NC；第9列写明每行匹配依据
- **人工复核**：GUI 内对未匹配行搜索候选料号、指定后自动重排重导出
- **生产版导出**：审核完（Excel 里手填料号、可追加无位号的 PCB 料号行）后一键清理——
  删 NC/非采购项/旧汇总行，按料号重新聚合排序，追加行置顶，
  输出 6 列（序号/料号/名称/规格/数量/位号）冻结表头的 `<原名>_生产版.xlsx`；漏填料号的行保留黄色置尾提醒
- **规则数据化**：替代/禁用规则可放 `rules.xlsx` 直接编辑，无需改代码

## 结构

```
app.py               pywebview 壳（gui/index.html 双皮肤）
engine.py            GUI 入口适配层（行回调/复核导出契约）
bom_matcher/         引擎内核：model / parsers / loader / rules / engine / report / finalize
engine_laoniu.py     laoniu（R系列编码表）引擎，独立保留
tests/               解析单测 + 金标准回归（tests/golden/，工具 make_golden.py）
BomMatcher.spec      PyInstaller 单文件打包（--add-data gui;gui）
```

## 使用

```powershell
# GUI
python app.py            # 或直接运行打包的 BomMatcher.exe

# 命令行
python engine.py --bom BOM.xlsx --code 物料编码表.xlsx [--output 结果.xlsx]

# 审核后生成生产版（GUI 里也有「生成生产版」按钮）
python engine.py --finalize 结果.xlsx [--output 生产版.xlsx]

# 测试（改引擎后必跑，金标准差异须逐条人工裁决）
python -m pytest tests/ -q
```

依赖：Python 3.8+，`openpyxl`、`pywebview`（GUI）、`pyinstaller`（打包）。Windows 10+（WebView2）。
