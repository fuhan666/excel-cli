# Excel-CLI

面向 AI、脚本和终端用户的 Excel 命令行工具。既提供 JSON API 供自动化调用，也内置了交互式 TUI，支持用类 Vim 快捷键浏览和编辑表格。

## 功能特性

- 用类 Vim 快捷键浏览、导航 Excel 工作表
- 在多表工作簿里创建、切换和删除工作表
- 直接在终端里编辑单元格内容
- 将数据导出为 JSON
- 支持列选择、筛选、分页和流式读取，方便自动化场景
- 对工作簿或单个工作表做质量检查，并以稳定的 JSON 格式输出结果
- 删除行和列
- 支持搜索并高亮匹配项
- 命令模式支持高级操作

## 安装与卸载

### 安装

#### 方式一：通过 Cargo 安装（推荐）

本包已发布到 crates.io，直接执行：

```bash
cargo install excel-cli --locked
```

#### 方式二：从 GitHub Release 下载

1. 访问 [GitHub Releases](https://github.com/fuhan666/excel-cli/releases)
2. 下载适合你系统的预编译二进制文件
3. 把可执行文件放到系统 PATH 里，或直接在下载目录运行

Linux 和 macOS 用户可能需要先添加执行权限。

#### 方式三：从源码编译

需要 Rust 和 Cargo。执行以下命令：

```bash
# 克隆仓库
git clone https://github.com/fuhan666/excel-cli.git
cd excel-cli
cargo build --release

# 安装到系统
cargo install --path . --locked
```

### 卸载

```bash
cargo uninstall excel-cli
```

## 使用方法

```bash
# 查看工作簿元数据
excel-cli inspect workbook path/to/your/file.xlsx

# 查看单个工作表
excel-cli inspect sheet path/to/your/file.xlsx --sheet Orders
excel-cli inspect sheet path/to/your/file.xlsx --sheet-index 0

# 从工作表采样数据
excel-cli inspect sample path/to/your/file.xlsx --sheet Orders --rows 10

# 查看列信息（自动检测表头）
excel-cli inspect columns path/to/your/file.xlsx --sheet Orders --header-row auto

# 检测表格区域
excel-cli inspect tables path/to/your/file.xlsx --sheet Orders

# 读取单个单元格
excel-cli read cell path/to/your/file.xlsx --sheet Orders --cell B2

# 读取区域
excel-cli read range path/to/your/file.xlsx --sheet Orders --range A1:F20

# 读取行（自动检测表头）
excel-cli read rows path/to/your/file.xlsx --sheet Orders

# 读取行并指定表头行（从 1 开始计数）
excel-cli read rows path/to/your/file.xlsx --sheet Orders --header-row 1

# 读取指定列作为记录
excel-cli read records path/to/your/file.xlsx --sheet Orders --select order_id,total,status

# 筛选、分页并以 JSON Lines 流式输出记录
excel-cli read records path/to/your/file.xlsx --sheet Orders \
  --filter status:eq:open \
  --filter total:gte:100 \
  --limit 50 \
  --output-shape jsonl

# 用完整规则集检查工作簿质量
excel-cli check path/to/your/file.xlsx

# 只检查一个工作表，并指定规则
excel-cli check path/to/your/file.xlsx --sheet Orders --rules duplicate_values,type_drift

# 只返回 warning 和 error 级别的结果
excel-cli check path/to/your/file.xlsx --severity-threshold warning

# 打开交互式 TUI 浏览器
excel-cli ui path/to/your/file.xlsx
```

### 命令行选项

所有非交互式命令（`inspect`、`read`、`check`）默认输出 JSON。加 `--format text` 可得到人类可读的文本输出。

**全局输出约定：**
- `stdout` 只输出结果
- `stderr` 只输出错误
- 成功时返回退出码 `0`
- 失败时返回非零退出码（见下方说明）
- 空单元格在 JSON 模式下输出 `null`，在文本模式下输出空字符串

### 读取行与记录

`read rows` 默认返回位置数组。加 `--output-shape records` 可返回以解析后的表头为键的对象；也可以直接用 `read records`，此时默认就是记录格式。

```bash
excel-cli read rows report.xlsx --sheet Orders --output-shape rows
excel-cli read rows report.xlsx --sheet Orders --output-shape records
excel-cli read records report.xlsx --sheet Orders
```

`--select` 用于保留指定列。列名来自解析后的表头行，重复或空白的表头会按与 `inspect columns` 相同的方式处理成稳定名称。

```bash
excel-cli read records report.xlsx --sheet Orders --select order_id,customer,total
```

`--filter 字段:操作符:值` 按列名筛选行。多次使用 `--filter` 会以 AND 逻辑组合条件。支持的操作符有 `eq`、`ne`、`gt`、`gte`、`lt`、`lte`、`contains`、`regex`、`isnull` 和 `notnull`。

```bash
excel-cli read records report.xlsx --sheet Orders --filter status:eq:open
excel-cli read records report.xlsx --sheet Orders --filter total:gte:100
excel-cli read records report.xlsx --sheet Orders --filter customer:contains:Inc
excel-cli read records report.xlsx --sheet Orders --filter order_id:regex:^INV-[0-9]+$
excel-cli read records report.xlsx --sheet Orders --filter optional_note:isnull:
```

`--limit` 和 `--offset` 在筛选之后生效。`--non-empty` 会去掉所有单元格都为空的行。即使筛选后没有匹配结果，也属于成功查询，返回空的 `rows` 或 `records` 数组，退出码为 `0`。

```bash
excel-cli read records report.xlsx --sheet Orders \
  --filter status:eq:open \
  --offset 25 \
  --limit 25 \
  --non-empty
```

`--output-shape jsonl` 将换行分隔的 JSON 记录直接输出到 stdout，而非标准包装格式。它沿用记录输出时的列选择、筛选、分页和表头解析规则。

```bash
excel-cli read records report.xlsx --sheet Orders --output-shape jsonl
```

如果列名不存在、筛选列未知、操作符不支持、筛选条件格式错误、数值比较无效或正则表达式无效，会返回结构化的 `invalid_query` 错误，退出码为 `6`。

### 质量检查

`check` 会对整个工作簿或单个工作表运行固定的质量规则集，输出格式与其他非交互式命令一致，采用稳定的 JSON 包装结构。默认扫描所有工作表，返回 `info`、`warning`、`error` 三级结果；过滤后仍有结果则退出码为 `1`，过滤后为空则退出码为 `0`。

支持的规则：
- `blank_headers`：标记检测到的表头行中的空白单元格
- `duplicate_headers`：标记标准化后重复的表头名称
- `blank_rows`：标记已用区域内的整行空白
- `blank_columns`：标记已用区域内的整列空白
- `null_ratio`：标记空值比例超过内置阈值的列
- `duplicate_values`：标记候选标识列中的重复值
- `type_drift`：标记同一列中偏离主类型的数据
- `formula_presence`：报告检查区域内仍包含公式的工作表

用 `--sheet <name>` 可限定只检查单个工作表，`--rules <逗号分隔的规则 ID>` 可按注册表顺序运行部分规则，`--severity-threshold <info|warning|error>` 可过滤返回的结果级别。`data.summary` 只统计最终返回的结果数，`data.stats.finding_count_before_threshold` 保留阈值过滤前的总数，`data.stats.rules_run` 记录规范化后的规则顺序。

```bash
excel-cli check report.xlsx
excel-cli check report.xlsx --sheet 客户 --rules blank_headers,duplicate_headers
excel-cli check report.xlsx --rules null_ratio,duplicate_values,type_drift --severity-threshold warning
```

### 退出码

| 代码 | 含义 |
|------|------|
| `0` | 成功 |
| `1` | 检查完成但发现问题 |
| `2` | 无效命令或参数 |
| `3` | 文件无法打开或读取 |
| `4` | 工作簿解析失败或格式不支持 |
| `5` | 未找到工作表、单元格、区域或目标 |
| `6` | 无效查询或检查规则 |
| `7` | 内部错误 |

### 输出格式

非交互式命令的成功响应采用统一的 JSON 包装格式：

```json
{
  "schema_version": "1.0",
  "command": "inspect.sheet",
  "file": { "path": "report.xlsx", "format": "xlsx" },
  "target": { "sheet": "Orders", "sheet_index": 1 },
  "meta": {},
  "data": { ... },
  "warnings": []
}
```

### 结构检查

`inspect columns` 分析工作表中的每一列，帮你为后续命令选择稳定的字段名。响应数据的 `columns` 数组中，每列包含 `index`、原始 `name`、生成的 `safe_name`、`is_duplicate`、推断的 `inferred_type`、`non_null_ratio`、`formula_ratio` 和 `sample_values`。响应元数据包含 `header_row_mode`、`resolved_header_row`、`column_count` 和 `data_row_count`。

```bash
excel-cli inspect columns path/to/your/file.xlsx --sheet Orders --header-row auto
excel-cli inspect columns path/to/your/file.xlsx --sheet Orders --header-row 2 --format text
```

`inspect tables` 检测工作表中的连续表格区域。响应数据的 `data.candidates` 中，每个候选区域包含 `range`、`header_row`、`column_count`、`row_count` 和 `confidence`。响应元数据包含 `candidate_count`。

```bash
excel-cli inspect tables path/to/your/file.xlsx --sheet Orders
excel-cli inspect tables path/to/your/file.xlsx --sheet Orders --format text
```

非交互式命令的错误响应也采用统一的 JSON 包装格式：

```json
{
  "schema_version": "1.0",
  "command": "read.rows",
  "file": { "path": "report.xlsx", "format": "xlsx" },
  "error": {
    "code": "target_not_found",
    "message": "Sheet 'Orders' not found",
    "details": {}
  }
}
```

## 用户界面

界面简洁直观：

- **标题栏与工作表标签**：显示当前文件名和所有可用工作表，当前工作表高亮显示
- **电子表格区域**：主数据展示区域
- **内容面板**：显示当前选中单元格的完整内容
- **通知面板**：显示操作反馈和系统通知
- **状态栏**：显示操作提示和当前输入的命令

## 键盘快捷键

- `h`、`j`、`k`、`l` 或方向键：在单元格之间移动（每次 1 格）
- `[`：切换到上一个工作表（停在第一个工作表）
- `]`：切换到下一个工作表（停在最后一个工作表）
- `0`：跳到当前行的第一列
- `^`：跳到当前行的第一个非空列
- `$`：跳到当前行的最后一列
- `gg`：跳到当前列的第一行
- `G`：跳到当前列的最后一行
- `Ctrl+←`（Mac 上为 `Command+←`）：当前单元格为空时跳到左侧第一个非空单元格；非空时跳到左侧最后一个非空单元格
- `Ctrl+→`（Mac 上为 `Command+→`）：当前单元格为空时跳到右侧第一个非空单元格；非空时跳到右侧最后一个非空单元格
- `Ctrl+↑`（Mac 上为 `Command+↑`）：当前单元格为空时跳到上方第一个非空单元格；非空时跳到上方最后一个非空单元格
- `Ctrl+↓`（Mac 上为 `Command+↓`）：当前单元格为空时跳到下方第一个非空单元格；非空时跳到下方最后一个非空单元格
- `Enter`：编辑当前单元格
- `y`：复制当前单元格内容
- `d`：剪切当前单元格内容
- `p`：将剪贴板内容粘贴到当前单元格
- `u`：撤销上一次操作（编辑、行列变更、工作表创建/删除）
- `Ctrl+r`：重做上一次撤销的操作
- `/`：开始向前搜索
- `?`：开始向后搜索
- `n`：跳到下一个搜索结果
- `N`：跳到上一个搜索结果
- `:`：进入命令模式（类 Vim 命令）

## Vim 编辑模式

编辑单元格内容时（按 `Enter` 进入编辑模式）：

- **模式切换**

  - `Esc`：退出 Vim 模式并保存修改
  - `i`：进入插入模式
  - `v`：进入可视模式

- **导航（普通模式）**

  - `h`、`j`、`k`、`l`：左、下、上、右移动光标
  - `w`：跳到下一个单词
  - `b`：跳到单词开头
  - `e`：跳到单词末尾
  - `$`：跳到行尾
  - `^`：跳到行首第一个非空白字符
  - `gg`：跳到第一行
  - `G`：跳到最后一行

- **编辑操作**

  - `x`：删除光标下的字符
  - `D`：删除到行尾
  - `C`：修改到行尾
  - `o`：在下方打开新行并进入插入模式
  - `O`：在上方打开新行并进入插入模式
  - `A`：在行尾追加
  - `I`：在行首插入

- **可视模式操作**

  - `y`：复制（yank）选中的文本
  - `d`：删除选中的文本
  - `c`：修改选中的文本（删除并进入插入模式）

- **操作符命令**

  - `y{motion}`：复制 motion 指定的文本
  - `d{motion}`：删除 motion 指定的文本
  - `c{motion}`：修改 motion 指定的文本

- **剪贴板操作**

  - `p`：粘贴复制或删除的文本
  - `u`：撤销上一次修改
  - `Ctrl+r`：重做上一次撤销的修改

## 搜索模式

按 `/`（向前搜索）或 `?`（向后搜索）进入搜索模式：

- 输入搜索关键词
- `Enter`：执行搜索并跳到第一个匹配项
- `Esc`：取消搜索
- `n`：跳到下一个匹配项（搜索执行后）
- `N`：跳到上一个匹配项（搜索执行后）
- 搜索结果以黄色高亮显示
- 搜索顺序为先逐行从左到右，再从上到下移动到下一行

## 命令模式

按 `:` 进入命令模式。可用命令如下：

### 列宽命令

- `:cw fit` — 自动调整当前列宽以适应内容
- `:cw fit all` — 自动调整所有列宽以适应内容
- `:cw min` — 最小化当前列宽（最大 15 或内容宽度）
- `:cw min all` — 最小化所有列宽（最大 15 或内容宽度）
- `:cw [数字]` — 将当前列宽设为指定值

### JSON 导出命令

- `:ej [h|v] [行数]` — 将当前工作表导出为 JSON

  - `h|v` — 表头方向：`h` 为横向（顶部行），`v` 为纵向（左侧列）
  - `行数` — 表头行数（横向）或列数（纵向）

- `:eja [h|v] [行数]` — 将所有工作表导出到单个 JSON 文件
  - 参数与 `:ej` 相同
  - 生成一个 JSON 对象，以工作表名为键，数据为值

输出文件名自动生成，格式如下：

- 单个工作表：`原文件名_sheet_工作表名称_YYYYMMDD_HHMMSS.json`
- 所有工作表：`原文件名_all_sheets_YYYYMMDD_HHMMSS.json`

JSON 文件保存在原始 Excel 文件所在目录。

### 类 Vim 命令

- `:w` — 保存文件但不退出
- `:wq` 或 `:x` — 保存并退出
- `:q` — 退出（如有未保存的修改会提示警告）
- `:q!` — 强制退出，不保存
  保存逻辑详见[下文](#文件保存逻辑)。

- `:y` — 复制当前单元格内容
- `:d` — 剪切当前单元格内容
- `:put` 或 `:pu` — 将剪贴板内容粘贴到当前单元格
- `:[单元格]` — 跳到指定单元格（如 `:A1`、`:B10`）。大小写不敏感（`:a1` 与 `:A1` 效果相同）

### 工作表管理命令

- `:addsheet [名称]` — 在当前工作表后添加新工作表
- `:sheet [名称/编号]` — 按名称或索引切换工作表（从 1 开始计数）
- `:delsheet` — 删除当前工作表

### 行列管理命令

- `:dr` — 删除当前行
- `:dr [行号]` — 删除指定行（如 `:dr 5` 删除第 5 行）
- `:dr [起始] [结束]` — 删除行范围（如 `:dr 5 10` 删除第 5 到 10 行）
- `:dc` — 删除当前列
- `:dc [列]` — 删除指定列（如 `:dc A`、`:dc a` 或 `:dc 1` 都删除 A 列）
- `:dc [起始] [结束]` — 删除列范围（如 `:dc A C` 或 `:dc a c` 删除 A 到 C 列）
- `:freeze` — 按当前单元格冻结其上方行和左侧列
- `:freeze [单元格]` — 按指定单元格冻结窗格（如 `:freeze B2` 冻结第 1 行和 A 列）
- `:unfreeze` — 取消当前工作表的冻结窗格

### 其他命令

- `:nohlsearch` 或 `:noh` — 关闭搜索高亮
- `:help` — 显示所有快捷键

## 文件保存逻辑

Excel-CLI 采用非破坏性保存方式：

- 保存时（使用 `:w`、`:wq` 或 `:x`），程序会检查是否有修改
- 如果没有修改，不会创建新文件，并提示 "No changes to save"
- 如果启用了懒加载，所有未加载的工作表会在保存前加载，以保留工作簿内容
- 如果有修改，会创建一个带时间戳的新文件，格式为 `原文件名_YYYYMMDD_HHMMSS.xlsx`
- 新文件不包含任何样式
- 原始文件永远不会被修改

## 贡献指南

分支命名、提交信息和 Pull Request 规范请参阅 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 许可证

MIT
