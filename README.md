# ship_tools

FBA 货件助手，用于把 Amazon 后台导出的装箱单、出货计划、货代模板和基础商品数据串起来，自动生成出货计划、货代发票/清关资料、标签裁剪文件和货件汇总。

## 功能概览

- `plan`: 根据 `plan_data/出货计划.xlsx` 生成 Amazon 批量货件模板和领星 ERP 出货计划模板。
- `invoice`: 根据 `plan_data/开票信息.xlsx` 和 `shipment_data/` 装箱单生成货代发票/交接资料，并处理 FBA 外箱标签。
- `invoice SEND`: 在 `invoice` 基础上，将 Amazon 外箱标签 PDF 按奇偶页拆分：奇数页为 FBA 外箱标，偶数页为唛头，并分别裁剪输出。
- `summary`: 汇总 `shipment_data/` 下所有装箱单的 SKU 数、盒数、箱数、体积和重量。
- 九方支持：当 `开票信息.xlsx` 中物流商为 `九方` 时，自动填充 `ship_partner/九方/` 下对应清关资料 Excel。

## 安装依赖

建议使用 Python 3，并在项目根目录安装依赖：

```bash
pip3 install -r requirements.txt
```

依赖包括：

- `pandas`: 读取出货计划 Excel。
- `openpyxl`: 读写 Excel 模板、插入图片。
- `PyMuPDF`: 处理 PDF 标签裁剪。
- `Pillow`: 读取图片并适配 Excel 单元格。

## 目录说明

```text
basic_data/
  data.txt                 全局商品/清关基础信息
  产品信息.csv              ASIN、FNSKU、SKU、套装类型、链接售卖价格
  产品包装信息.csv           套装类型对应箱规、价格、重量、图片
  picture/                 商品图片
  template/invoice/         货代发票模板
  template/plan/            Amazon 和领星 ERP 出货计划模板

plan_data/
  出货计划.xlsx              plan 命令输入
  开票信息.xlsx              invoice 命令输入

shipment_data/
  *.csv                    Amazon 后台导出的装箱单
  *.pdf                    Amazon 生成的外箱标签 PDF

ship_partner/
  九方/*.xlsx               九方清关资料 Excel，文件名需以货件 ID 结尾

output/
  yyyyMMdd/                所有生成结果按日期输出到这里

src/
  invoice/                 发票、九方资料填充逻辑
  plan/                    出货计划生成逻辑
  fba_label/               FBA 外箱标和唛头 PDF 处理逻辑
  summary/                 货件汇总逻辑
```

`shipment_data/` 和 `output/` 已加入 `.gitignore`，适合放临时输入和生成结果。

## 基础数据说明

### `basic_data/data.txt`

按 `key=value` 格式维护全局字段。当前主流程会使用：

```text
Product_Name_Cn=毛巾
Product_Name_En=Towel
HS_Code=6302910000
Material=纯棉/Cotton
Usr_For=清洁/Clean
Brand=SEMAXE
Model=None
Unit=盒
Product_Attribute=普货（无任何电池）no battery
```

### `basic_data/产品信息.csv`

以 ASIN 为核心维护商品基础信息。常用字段：

- `颜色`
- `套装类型`
- `ASIN`
- `FNSKU`
- `SKU`
- `链接售卖价格`

九方资料填充时会根据 Excel 中的 `ASIN Code` 查这里，并读取 `链接售卖价格`。

### `basic_data/产品包装信息.csv`

以 `套装类型` 为 key，维护箱规、申报信息和图片：

- `单箱数量`
- `长(cm)`、`宽(cm)`、`高(cm)`
- `采购单价`
- `申报单价`
- `单产品重量(kg)`
- `净重(kg)`
- `毛重(kg)`
- `图片`

发票、计划、summary、九方资料填充都会复用这里的数据。

## 命令说明

所有命令均在项目根目录执行。

### 查看帮助

```bash
python3 ship_tool.py --help
```

### 生成出货计划

```bash
python3 ship_tool.py plan
```

输入：

- `plan_data/出货计划.xlsx`
- `basic_data/产品信息.csv`
- `basic_data/产品包装信息.csv`
- `basic_data/template/plan/*.xlsx`

处理逻辑：

- 读取 `出货计划.xlsx` 第一个 sheet。
- sheet 名最后一段用于识别国家，例如 `SEMAXE-US` 会识别为 `US`。
- 按颜色和套装类型查找 SKU/FNSKU。
- 根据国家选择 Amazon US 或 CA 模板。
- 同时生成领星 ERP 出货计划。

输出：

```text
output/yyyyMMdd/plan/Amazon_Shipment_US_yyyyMMdd.xlsx
output/yyyyMMdd/plan/ERP_Shipment_Plan_yyyyMMdd.xlsx
```

## 发票和标签处理

### 普通 invoice 模式

```bash
python3 ship_tool.py invoice
```

等价于：

```bash
python3 ship_tool.py invoice FIST
```

输入：

- `plan_data/开票信息.xlsx`
- `shipment_data/*.csv`
- `shipment_data/*.pdf`
- `basic_data/data.txt`
- `basic_data/产品信息.csv`
- `basic_data/产品包装信息.csv`
- `basic_data/template/invoice/*.xlsx`

处理逻辑：

- 读取 `开票信息.xlsx`，逐行按物流商处理。
- 解析 `shipment_data/` 下全部 Amazon 装箱单 CSV。
- 根据物流商调用对应模板填充逻辑。
- 每个仓库发票生成后输出当前仓库汇总。
- 处理 `shipment_data/` 中的 FBA 外箱标签 PDF，执行当前裁剪逻辑。

FIST 模式下标签输出：

```text
output/yyyyMMdd/fba_label/原PDF文件名.pdf
```

### SEND 模式

```bash
python3 ship_tool.py invoice SEND
```

SEND 模式适用于 Amazon 生成的 PDF 中奇偶页混合的场景：

- 奇数页：FBA 外箱标。
- 偶数页：唛头。

处理逻辑：

- 奇数页提取后继续执行 FBA 外箱标裁剪和擦除目的地公司名称。
- 偶数页提取后裁掉空白，并适配到 `100mm x 80mm` 标签纸。

输出：

```text
output/yyyyMMdd/fba_label/FBA标签_货件ID_箱数.pdf
output/yyyyMMdd/唛头/唛头_货件ID_箱数.pdf
```

示例：

```text
shipment_data/FBA19D6YR2MZ-1779693385368.pdf
```

SEND 模式会输出：

```text
output/yyyyMMdd/fba_label/FBA标签_FBA19D6YR2MZ_51.pdf
output/yyyyMMdd/唛头/唛头_FBA19D6YR2MZ_51.pdf
```

## 支持的货代

当前 `invoice` 主流程支持以下物流商：

- `紐酷`
- `赤道`
- `鹏城海`
- `易通安达`
- `壹港`
- `百利通`
- `九方`

物流商名称需要填写在 `plan_data/开票信息.xlsx` 的 `物流商` 列中。

## 九方资料填充

当 `plan_data/开票信息.xlsx` 中某一行的物流商为 `九方` 时，主流程会执行九方清关资料填充逻辑。

### 输入要求

九方待填充文件放在：

```text
ship_partner/九方/
```

文件名必须以货件 ID 结尾：

```text
*FBA19XXXXXXX.xlsx
```

同时，Excel 内部 `C4` 单元格也必须等于货件 ID。

### 前置校验

在正式处理发票前，系统会先校验所有九方记录：

- 找不到 `*货件ID.xlsx`: 输出 warning。
- 找到多个匹配文件: 输出 warning。
- Excel 内部 `C4` 和货件 ID 不一致: 输出 warning。

所有九方记录校验完成后，如果存在任一问题，会一次性终止 `invoice`。

### 字段来源

九方 Excel 填充字段：

- `D列 Reference ID`: `开票信息.xlsx` 的 `追踪编号`
- `G列 英文品名`: `basic_data/data.txt` 的 `Product_Name_En`
- `H列 中文品名`: `Product_Name_Cn`
- `I列 品牌`: `Brand`
- `J列 材质`: `Material`
- `K列 用途`: `Usr_For`
- `L列 型号`: `Model`
- `M列 HS CODE`: `HS_Code`
- `O列 单位`: `Unit`
- `P列 每套个数`: `产品包装信息.csv` 的 `单箱数量`
- `Q列 采购单价`: `产品包装信息.csv` 的 `采购单价`
- `R列 申报单价`: `产品包装信息.csv` 的 `申报单价`
- `S列 链接售卖价格`: `产品信息.csv` 的 `链接售卖价格`
- `U列 产品属性`: `Product_Attribute`
- `V列 图片`: `产品包装信息.csv` 的 `图片`

如果填充阶段发现 ASIN 找不到、套装类型找不到、链接售卖价格为空或图片不存在，会直接终止程序。

输出：

```text
output/yyyyMMdd/invoice/原文件名.xlsx
```

成功日志示例：

```text
【FBA19xxxxx】-【仓库代码】开具[九方]发票成功...
```

## 汇总功能

```bash
python3 ship_tool.py summary
```

输入：

- `shipment_data/*.csv`
- `basic_data/产品信息.csv`
- `basic_data/产品包装信息.csv`

输出日志示例：

```text
【FBA19D70W3D0】-【CLT2】: {'总sku数': 36, '总盒数': 704, '总箱数': 57, '总体积(m³)': 4.7254, '总重量(kg)': 1006.16}
```

汇总口径：

- `总sku数`: 装箱单明细中唯一 SKU 数。
- `总盒数`: 装箱单商品总数合计。
- `总箱数`: 装箱单箱数合计。
- `总体积(m³)`: 按 `产品包装信息.csv` 的长宽高和箱数计算。
- `总重量(kg)`: 按 `产品包装信息.csv` 的毛重和箱数计算。

`invoice` 命令中每个仓库发票生成成功后，也会自动输出同样口径的汇总。

## 开票信息文件格式

`plan_data/开票信息.xlsx` 的 `开票信息` sheet 从第 3 行开始读取数据。当前使用列：

- `货件ID`
- `追踪编号`
- `是否退税`
- `物流商`
- `渠道类型`
- `送货时间`
- `预约船期`
- `货件创建时间`
- `仓库代码`

基础校验：

- 前 4 列必填：货件 ID、追踪编号、是否退税、物流商。
- `赤道`: 所有字段必填。
- `紐酷`: 渠道类型和送货时间必填。

## 装箱单要求

`shipment_data/` 下的 CSV 应为 Amazon 后台导出的装箱单。

解析逻辑依赖：

- 前 9 行为货件基础信息。
- 第 10 行为明细表头。
- 明细中会读取 SKU、ASIN、箱子总数、商品总数和箱号。

如果 ASIN 无法在 `basic_data/产品信息.csv` 中找到，会终止处理。

## 常见问题

### 运行 plan 或 invoice 找不到文件

确认以下目录和文件存在：

```text
basic_data/
plan_data/
shipment_data/
ship_partner/九方/   # 仅九方需要
```

### 九方资料校验失败

检查：

- `开票信息.xlsx` 中物流商是否为 `九方`。
- `ship_partner/九方/` 下是否存在 `*货件ID.xlsx`。
- 是否有多个文件同时以同一货件 ID 结尾。
- Excel 内部 `C4` 是否等于货件 ID。

### 九方填充时报链接售卖价格为空

补充 `basic_data/产品信息.csv` 中对应 ASIN 的 `链接售卖价格`。

### 标签输出不是预期尺寸

- FBA 外箱标输出使用既有裁剪区域。
- SEND 模式唛头输出会适配到 `100mm x 80mm`。

## 开发说明

主要入口：

- [ship_tool.py](ship_tool.py)

主要模块：

- `src/plan/process_ship_plan.py`: 出货计划生成。
- `src/invoice/process_ship_invoice.py`: invoice 主流程。
- `src/invoice/process_chidao.py`: 赤道特殊处理。
- `src/invoice/process_jiufang.py`: 九方资料填充。
- `src/fba_label/process_fba_label.py`: FBA 外箱标和唛头 PDF 处理。
- `src/summary/process_summary.py`: 货件汇总。
- `src/utils/common_utils.py`: 基础数据加载、日志、日期和输出目录。
- `src/utils/excel_utils.py`: Excel 图片、样式等工具函数。
