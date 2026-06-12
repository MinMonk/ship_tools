# Agent 行为规范

本文档用于约束后续在本项目中协作的 AI Agent 行为。目标是让每次修改都可预期、可验证，并尽量贴合当前项目已经形成的业务习惯。

## 协作原则

1. 任何代码修改、配置修改、数据结构调整前，必须先输出对需求的理解、实现思路和执行计划。
2. 当用户明确说“先澄清”“不要动代码”“先不要改代码”时，只允许阅读、扫描、分析和提问，不允许修改任何文件。
3. 需求存在歧义时，优先列出不确定点继续澄清；需求已经明确时，直接按计划执行并完成验证。
4. 修改应保持小范围、可回溯，优先沿用当前项目已有目录结构、命名风格、日志风格和处理模式。
5. 工作区可能存在用户未提交的改动，不得回退、覆盖或清理非本次修改产生的内容。
6. 不使用破坏性命令，例如 `git reset --hard`、强制删除业务文件等；如确需删除文件，必须确认这是用户明确要求。
7. 最终回复必须说明改了什么、验证了什么、还有哪些风险或未验证项。

## 语言与沟通

1. 默认使用中文回复。
2. 中间进度说明保持简短，说明正在查看什么、学到了什么、下一步做什么。
3. 对业务规则要按用户给出的口径执行，不擅自改变默认行为。
4. 用户给出具体输出格式、文件名、目录或日志文本时，应尽量逐字遵守。

## 工具与编辑习惯

1. 查找文件和文本优先使用 `rg`、`rg --files`。
2. 手工编辑文件必须使用 `apply_patch`。
3. 不用 `cat > file`、shell 重定向、临时 Python 脚本等方式进行普通代码编辑。
4. 可以使用 Python 或表格库读取、校验、迁移 Excel/CSV 数据，但执行前要说明目的。
5. 修改后优先运行与改动相关的最小验证命令，例如：
   - `python3 -m compileall ship_tool.py src`
   - `python3 ship_tool.py plan`
   - `python3 ship_tool.py invoice`
   - `python3 ship_tool.py summary`
   - `python3 ship_tool.py calc`
6. 如果验证失败，要区分代码问题、数据问题、环境问题，并把关键错误说明清楚。

## 项目入口与命令

项目主入口是 `ship_tool.py`。用户偶尔会写成 `ship_tools.py`，实际执行时应以当前项目存在的 `ship_tool.py` 为准。

当前主要命令：

```bash
python3 ship_tool.py plan [US|CA]
python3 ship_tool.py invoice
python3 ship_tool.py summary
python3 ship_tool.py calc
```

## 目录约定

1. `basic_data/`：基础配置、产品信息、包装信息、模板和图片。
2. `plan_data/`：出货计划和开票信息。
3. `shipment_data/`：物流模式输入目录和运费价格表。
4. `shipment_data/FIST/`：FIST 模式装箱单和外箱标签。
5. `shipment_data/SEND/`：SEND 模式装箱单和外箱标签。
6. `shipment_data/ship_price.xlsx`：海运费测算价格表。
7. `ship_partner/九方/`：九方货代资料模板来源。
8. `output/yyyyMMdd/`：运行生成的输出文件。

注意：`shipment_data/FIST/` 和 `shipment_data/SEND/` 下的业务输入文件不提交到 Git；`shipment_data/ship_price.xlsx` 可以作为项目数据提交。

## 现有业务规则

### plan

1. `python3 ship_tool.py plan` 不带参数时，遍历《出货计划.xlsx》所有有效 sheet。
2. `python3 ship_tool.py plan CA` 或 `US` 时，只处理指定国家。
3. 只处理 sheet 名以 `-US` 或 `-CA` 结尾的发货计划 sheet。
4. sheet 名不符合规则时输出 warning 并跳过。
5. 带国家参数但没有对应计划时，提示 `【CA】该国家没有出货计划` 这类信息。
6. US 模板使用 inch/lb；CA 模板保持 cm/kg。
7. CA 模板顶部默认值填写 `Seller`，明细行 C/D 列保持为空。
8. 当前 ERP 模板生成功能处于注释或停用状态，除非用户重新要求，不要恢复。

### invoice

1. `invoice` 不再接收命令行第二参数。
2. 全局物流模式来自 `plan_data/开票信息.xlsx` 的 `G1` 单元格，只允许 `FIST` 或 `SEND`。
3. `G1` 为空时提示请选择物流模式并终止。
4. FIST 模式读取 `shipment_data/FIST/`，SEND 模式读取 `shipment_data/SEND/`。
5. 对应模式目录不存在或没有 CSV 时终止；没有 PDF 时输出 warning。
6. SEND 模式外箱标签处理逻辑：
   - 奇数页为 FBA 标签，继续执行当前 FBA 裁剪逻辑。
   - 偶数页为唛头标签，执行唛头裁剪处理。
7. 输出路径保持在 `output/yyyyMMdd/` 下对应子目录。
8. 日志输出属于用户体验的一部分，新增流程要补齐开始、模式、操作、完成等阶段日志。

### summary

1. 同时统计 `shipment_data/FIST/` 和 `shipment_data/SEND/`。
2. 输出分为 FIST 和 SEND 两部分。
3. 统计 key 使用 `物流模式-仓库名`。
4. FIST 的仓库名按装箱单中的当前逻辑获取。
5. SEND 的仓库名优先使用 CSV 文件名。
6. 目录无数据时照常输出 `暂无数据`。

### calc

1. 读取 `shipment_data/ship_price.xlsx` 中的 `FIST` 和 `SEND` sheet。
2. 费用基于 summary 统计出的重量计算，重量单位为 kg，单价默认单位为 元/kg。
3. FIST 模式允许不同仓库使用不同承运商，按每个仓库最终费用最低选择：
   - 最终费用 = 仓库重量 * 单价 + 承运商报关费。
   - 若多个承运商最终费用相同，全部输出。
   - 某仓库无任何承运商报价时 warning 后跳过。
4. SEND 模式要求所有仓库使用同一承运商：
   - 总费用 = sum(各仓库重量 * 各仓库单价 + 报关费)。
   - 承运商缺少任意一个仓库报价时跳过，并 warning。
   - 输出按总费用从低到高排序。
5. 分别测算 FIST 和 SEND 后，要输出最终建议方案，只比较总费用最低。

### 货代处理

1. 货代常量集中在 `src/constants/carriers.py`。
2. 九方只在《开票信息.xlsx》物流商为 `九方` 时执行：
   - 来源目录固定为 `ship_partner/九方/`。
   - 文件名匹配规则为以货件 ID 结尾的 `.xlsx`。
   - 前置校验要一次性检查缺失、多文件匹配、Excel 内部 C4 不匹配。
   - 输出到 `output/yyyyMMdd/invoice/`。
3. 联宇逻辑应独立在 `src/invoice/process_lianyu.py`：
   - 模板为 `basic_data/template/invoice/联宇.xlsx`。
   - 不需要图片。
   - 按仓库生成文件，每个文件只保留一个 sheet。
   - 输出文件名为 `联宇_仓库代码_总箱数_yyyyMMdd.xlsx`。

## README 与文档

1. 新增或调整功能后，如涉及用户命令、目录结构、输入文件、输出文件或注意事项，应同步考虑是否需要更新 `README.md`。
2. 文档应写清楚命令、输入数据、输出路径、异常处理和典型使用方式。

## 最终回复格式

完成修改后，最终回复应包含：

1. 修改内容摘要。
2. 关键文件路径。
3. 已执行的验证命令和结果。
4. 若未运行验证，说明原因。
5. 若存在数据依赖、未覆盖场景或后续建议，应明确列出。
