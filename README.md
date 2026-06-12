# ship_tools

FBA 货件助手，用于把 Amazon 后台导出的装箱单、出货计划、货代模板和基础商品数据串起来，自动生成出货计划、货代发票/清关资料、标签裁剪文件和货件汇总。

## 功能概览

- `plan`: 根据 `plan_data/出货计划.xlsx` 生成 Amazon 批量货件模板。
- `invoice`: 根据 `plan_data/开票信息.xlsx` 和 `shipment_data/` 装箱单生成货代发票/交接资料，并处理 FBA 外箱标签。
- `label`: 独立裁剪 `shipment_data/FIST` 或 `shipment_data/SEND` 下的标签 PDF，不依赖 `invoice`。
- `summary`: 汇总 `shipment_data/` 下所有装箱单的 SKU 数、盒数、箱数、体积和重量。
- `calc`: 根据 `summary` 重量和 `shipment_data/ship_price.xlsx` 测算 FIST/SEND 海运费。

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
  ship_price.xlsx          FIST/SEND 承运商报价表
  FIST/                    FIST 模式输入目录
    *.csv                  Amazon 后台导出的装箱单
    *.pdf                  Amazon 生成的外箱标签 PDF
  SEND/                    SEND 模式输入目录
    *.csv                  Amazon 后台导出的装箱单
    *.pdf                  Amazon 生成的外箱标签 PDF

ship_partner/
  九方/*.xlsx               九方清关资料 Excel，文件名需以货件 ID 结尾

output/
  yyyyMMdd/                所有生成结果按日期输出到这里

src/
  invoice/                 发票资料填充逻辑
  price/                   海运费测算逻辑
  plan/                    出货计划生成逻辑
  fba_label/               FBA 外箱标和唛头 PDF 处理逻辑
  summary/                 货件汇总逻辑
```

`output/` 已加入 `.gitignore`，适合放生成结果。`shipment_data/FIST/` 和 `shipment_data/SEND/` 会被忽略，`shipment_data/ship_price.xlsx` 可提交。

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
python3 ship_tool.py plan US
python3 ship_tool.py plan CA
```

输入：

- `plan_data/出货计划.xlsx`
- `basic_data/产品信息.csv`
- `basic_data/产品包装信息.csv`
- `basic_data/template/plan/*.xlsx`

处理逻辑：

- 不带国家参数时，遍历 `出货计划.xlsx` 所有 sheet。
- 带 `US` 或 `CA` 参数时，只处理对应国家的 sheet。
- 只处理以 `-US` 或 `-CA` 结尾的 sheet，例如 `SEMAXE-US` 会识别为 `US`。
- 非 `US/CA` 结尾的 sheet 会输出 warning 并跳过。
- 如果匹配到的国家 sheet 没有任何发货计划，会输出 `【国家】该国家没有出货计划`。
- 按颜色和套装类型查找 SKU/FNSKU。
- 根据国家选择 Amazon US 或 CA 模板。
- 当前暂不生成领星 ERP 出货计划。
- US 模板使用 inch/lb，CA 模板使用 cm/kg。
- CA 模板顶部默认值 `Default prep owner` 和 `Default labeling owner` 填写 `Seller`，明细行 C/D 列保持为空。

输出：

```text
output/yyyyMMdd/plan/Amazon_Shipment_US_yyyyMMdd.xlsx
output/yyyyMMdd/plan/Amazon_Shipment_CA_yyyyMMdd.xlsx
```

## 发票和标签处理

### invoice 模式

```bash
python3 ship_tool.py invoice
```

输入：

- `plan_data/开票信息.xlsx`
- `shipment_data/FIST/*.csv` 或 `shipment_data/SEND/*.csv`
- `shipment_data/FIST/*.pdf` 或 `shipment_data/SEND/*.pdf`
- `basic_data/data.txt`
- `basic_data/产品信息.csv`
- `basic_data/产品包装信息.csv`
- `basic_data/template/invoice/*.xlsx`

处理逻辑：

- 校验 `开票信息.xlsx` 的 `G1` 物流模式，只允许 `FIST` 或 `SEND`。
- 根据 `G1` 选择输入目录：`shipment_data/FIST` 或 `shipment_data/SEND`。
- 读取 `开票信息.xlsx`，逐行按物流商处理。
- 解析当前模式目录下全部 Amazon 装箱单 CSV。
- 根据物流商调用对应模板填充逻辑。
- 每个仓库发票生成后输出当前仓库汇总。

`G1 = FIST` 时，处理 `shipment_data/FIST/` 中的 FBA 外箱标签 PDF，执行当前裁剪逻辑。

```text
output/yyyyMMdd/fba_label/原PDF文件名.pdf
```

`G1 = SEND` 时，处理 `shipment_data/SEND/` 中的 PDF，适用于 Amazon 生成的 PDF 中奇偶页混合的场景：

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
shipment_data/SEND/FBA19D6YR2MZ-1779693385368.pdf
```

SEND 模式会输出：

```text
output/yyyyMMdd/fba_label/FBA标签_FBA19D6YR2MZ_51.pdf
output/yyyyMMdd/唛头/唛头_FBA19D6YR2MZ_51.pdf
```

### 独立裁剪标签

如果只需要重新裁剪标签，不需要重新生成发票，可以直接执行：

```bash
python3 ship_tool.py label FIST
python3 ship_tool.py label SEND
```

参数说明：

- `FIST` 或 `SEND` 大小写不敏感，例如 `send` 会按 `SEND` 处理。
- `label FIST` 读取 `shipment_data/FIST/*.pdf`，只裁剪 FBA 外箱标。
- `label SEND` 读取 `shipment_data/SEND/*.pdf`，一次性执行两段：
  - 奇数页拆分为 FBA 外箱标，并执行现有 FBA 裁剪和擦除逻辑。
  - 偶数页拆分为唛头，并裁掉空白、适配到 `100mm x 80mm` 标签纸。

校验规则：

- 对应模式目录不存在时终止。
- 对应模式目录没有 PDF 时终止。
- `label` 命令不校验 CSV 装箱单。
- 输出路径和覆盖规则沿用 `invoice` 后置标签处理逻辑。

## 支持的货代

当前 `invoice` 主流程支持以下物流商：

- `紐酷`
- `赤道`
- `鹏城海`
- `易通安达`
- `壹港`
- `百利通`
- `九方`
- `联宇`

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

## 汇总功能

```bash
python3 ship_tool.py summary
```

输入：

- `shipment_data/FIST/*.csv`
- `shipment_data/SEND/*.csv`
- `basic_data/产品信息.csv`
- `basic_data/产品包装信息.csv`

输出日志示例：

```text
▶ FIST 汇总
【FIST-TEB9】: {'总sku数': 31, '总盒数': 700, '总箱数': 51, '总体积(m³)': 4.3267, '总重量(kg)': 905.33}

▶ SEND 汇总
【SEND-TEB9】: {'总sku数': 31, '总盒数': 700, '总箱数': 51, '总体积(m³)': 4.3267, '总重量(kg)': 905.33}
```

汇总口径：

- 按 `FIST` 和 `SEND` 两部分分别输出。
- 输出 key 统一为 `物流模式-仓库名`。
- `FIST` 使用装箱单中的仓库名。
- `SEND` 使用 CSV 文件名作为仓库名，适配货件编号和货件名称为“暂未分配”的半成品装箱单。
- 任一模式目录不存在或没有 CSV 时，该模式输出 `暂无数据`。
- `总sku数`: 装箱单明细中唯一 SKU 数。
- `总盒数`: 装箱单商品总数合计。
- `总箱数`: 装箱单箱数合计。
- `总体积(m³)`: 按 `产品包装信息.csv` 的长宽高和箱数计算。
- `总重量(kg)`: 按 `产品包装信息.csv` 的毛重和箱数计算。

`invoice` 命令中每个仓库发票生成成功后，也会自动输出同样口径的汇总。

## 海运费测算

```bash
python3 ship_tool.py calc
```

输入：

- `shipment_data/ship_price.xlsx`
- `shipment_data/FIST/*.csv`
- `shipment_data/SEND/*.csv`
- `basic_data/产品信息.csv`
- `basic_data/产品包装信息.csv`

报价表结构：

- `FIST` 和 `SEND` 两个 sheet 必须存在。
- 第 1 行 B 列开始为承运商名称。
- 第 2 行 B 列开始为承运商报关费。
- 第 3 行开始 A 列为仓库名，B 列开始为各承运商仓库单价。
- 单价单位默认为元/kg。

测算规则：

- `FIST`: 每个仓库可选择不同承运商，按 `重量 * 仓库单价 + 报关费` 选择最低费用。
- `SEND`: 所有仓库必须选择同一个承运商，按 `sum(各仓库重量 * 仓库单价 + 报关费)` 计算总费用。
- 如果承运商缺少任意一个 SEND 仓库报价，则跳过该承运商并输出 warning。
- 如果 FIST 某个仓库没有任何承运商报价，则跳过该仓库并输出 warning。

输出日志示例：

```text
▶ FIST 海运费测算
【FIST-TEB9】承运商: 联宇，总箱数： 51 箱，重量: 905.33 kg，单价: 3.6 元/kg，报关费: 100元，费用: 3359.19元
FIST 预计总费用: 31234.56元

▶ SEND 海运费测算
【联宇】预计总费用: 30123.45元
    【TEB9】总箱数： 51 箱，重量: 905.33 kg，单价: 3.6 元/kg，报关费: 100元，费用: 3359.19元
```

## 开票信息文件格式

`plan_data/开票信息.xlsx` 的 `开票信息` sheet 使用 `G1` 作为整批物流模式，允许值为 `FIST` 或 `SEND`。如果 `G1` 为空或不是这两个值，`invoice` 会终止。

明细数据从第 3 行开始读取。当前使用列：

- `货件ID`
- `追踪编号`
- `仓库代码`
- `是否退税`
- `物流商`
- `渠道类型`
- `送货时间`
- `预约船期`
- `货件创建时间`

基础校验：

- 基础必填：货件 ID、追踪编号、仓库代码、是否退税、物流商。
- `赤道`: 所有业务字段必填。
- `紐酷`: 渠道类型和送货时间必填。

## 装箱单要求

`shipment_data/FIST/` 和 `shipment_data/SEND/` 下的 CSV 应为 Amazon 后台导出的装箱单。`invoice` 会按 `开票信息.xlsx` 的 `G1` 物流模式只读取对应目录。

校验规则：

- 对应模式目录不存在时终止。
- 对应模式目录没有 CSV 时终止。
- 对应模式目录没有 PDF 时打印 warning，发票仍会继续生成。

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
  FIST/
  SEND/
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
- `src/invoice/process_lianyu.py`: 联宇发票填充。
- `src/fba_label/process_fba_label.py`: FBA 外箱标和唛头 PDF 处理。
- `src/summary/process_summary.py`: 货件汇总。
- `src/utils/common_utils.py`: 基础数据加载、日志、日期和输出目录。
- `src/utils/excel_utils.py`: Excel 图片、样式等工具函数。
