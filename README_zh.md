# Excel-CLI

面向 AI、脚本和终端用户的 Excel 命令行工具。提供 JSON API 供自动化场景调用，同时内置交互式 TUI，支持类 Vim 快捷键浏览和编辑表格。

## 功能特性

- 使用类 Vim 快捷键浏览和导航 Excel 工作表
- 在多工作表工作簿中创建、切换和删除工作表
- 直接在终端中编辑单元格内容
- 将数据导出为 JSON 格式
- 选择、筛选、分页和流式读取结果，支持自动化场景
- 对整个工作簿或单个工作表运行质量检查，并输出稳定的 JSON findings
- 删除行和列
- 搜索功能并支持高亮显示
- TUI 中只读查询预览
- 在 TUI 中查看质量检查 findings
- 命令模式支持高级操作

## 安装与卸载

### 安装

#### 方式一：通过 Cargo 安装（推荐）

本包已发布到 crates.io，可直接使用以下命令安装：

```bash
cargo install excel-cli --locked
```

#### 方式二：从 GitHub Release 下载

1. 访问 [GitHub Releases](https://github.com/fuhan666/excel-cli/releases)
2. 下载适合您操作系统的预编译二进制文件
3. 将可执行文件放入系统路径，或直接从下载位置运行

Linux 和 macOS 用户可能需要先添加执行权限

#### 方式三：从源码编译

需要 Rust 和 Cargo。使用以下命令安装：

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
# 检查工作簿元数据
excel-cli inspect workbook path/to/your/file.xlsx

# 检查单个工作表
excel-cli inspect sheet path/to/your/file.xlsx --sheet Orders
excel-cli inspect sheet path/to/your/file.xlsx --sheet-index 0

# 从工作表中采样数据
excel-cli inspect sample path/to/your/file.xlsx --sheet Orders --rows 10

# 检查列信息（自动检测表头）
excel-cli inspect columns path/to/your/file.xlsx --sheet Orders --header-row auto

# 检查表格区域
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

# 使用完整 v1.3.0 规则集检查工作簿质量
excel-cli check path/to/your/file.xlsx

# 仅检查一个工作表，并选择指定规则
excel-cli check path/to/your/file.xlsx --sheet Orders --rules duplicate_values,type_drift

# 只返回 warning 和 error 级别的 findings
excel-cli check path/to/your/file.xlsx --severity-threshold warning

# 打开交互式 TUI 浏览器
excel-cli ui path/to/your/file.xlsx
```

### 命令行选项

所有无界面命令（`inspect`、`read`、`check`）默认输出 JSON 格式。使用 `--format text` 获取人类可读的输出。

**全局输出规则：**
- `stdout` 仅包含结果
- `stderr` 仅包含错误
- 成功返回退出码 `0`
- 失败返回非零退出码（见下方退出码说明）
- 空单元格在 JSON 模式下输出 `null`，在文本模式下输出空字符串

### 读取行与记录

`read rows` 默认返回位置数组。使用 `--output-shape records` 可以返回以解析后的表头为键的对象，或者直接使用 `read records`，此时记录形式输出为默认。

```bash
excel-cli read rows report.xlsx --sheet Orders --output-shape rows
excel-cli read rows report.xlsx --sheet Orders --output-shape records
excel-cli read records report.xlsx --sheet Orders
```

`--select` 保留指定的列，列名来自解析后的表头行，重复或空白的表头会按与 `inspect columns` 相同的方式处理为稳定名称。

```bash
excel-cli read records report.xlsx --sheet Orders --select order_id,customer,total
```

`--filter 字段:操作符:值` 按列名筛选行。重复 `--filter` 会以 AND 逻辑组合条件。支持的操作符有 `eq`、`ne`、`gt`、`gte`、`lt`、`lte`、`contains`、`regex`、`isnull` 和 `notnull`。

```bash
excel-cli read records report.xlsx --sheet Orders --filter status:eq:open
excel-cli read records report.xlsx --sheet Orders --filter total:gte:100
excel-cli read records report.xlsx --sheet Orders --filter customer:contains:Inc
excel-cli read records report.xlsx --sheet Orders --filter order_id:regex:^INV-[0-9]+$
excel-cli read records report.xlsx --sheet Orders --filter optional_note:isnull:
```

`--limit` 和 `--offset` 在筛选后生效。`--non-empty` 会移除所有单元格为空的行。未匹配的筛选仍是成功的查询：返回空的 `rows` 或 `records` 数组，退出码为 `0`。

```bash
excel-cli read records report.xlsx --sheet Orders \
  --filter status:eq:open \
  --offset 25 \
  --limit 25 \
  --non-empty
```

`--output-shape jsonl` 将换行分隔的 JSON 记录直接输出到 stdout，而不是标准的信封格式。它使用与记录输出相同的选择、筛选、分页和表头解析规则。

```bash
excel-cli read records report.xlsx --sheet Orders --output-shape jsonl
```

无效的选择列、未知的筛选列、不支持的操作符、格式错误的筛选条件、无效的数值比较和无效的正则表达式会返回结构化的 `invalid_query` 错误，退出码为 `6`。

### 质量检查

`check` 会对整个工作簿或单个工作表运行固定的 v1.3.0 质量规则集，并返回与其他无界面命令一致的稳定 JSON 信封结构。默认会扫描所有工作表，返回 `info`、`warning`、`error` 三种级别的 findings；当过滤后的 findings 非空时退出码为 `1`，过滤后为空时退出码为 `0`。

支持的规则：
- `blank_headers`：标记检测到的表头行中的空白表头单元格
- `duplicate_headers`：标记标准化后重复的表头名称
- `blank_rows`：标记已使用区域中的整空白行
- `blank_columns`：标记已使用区域中的整空白列
- `null_ratio`：根据内置阈值标记空值比例过高的列
- `duplicate_values`：标记候选标识列中的重复值
- `type_drift`：标记同一列中偏离主类型的数据
- `formula_presence`：报告检查区域中仍然包含公式的工作表

使用 `--sheet <name>` 可以只检查单个工作表，`--rules <以逗号分隔的规则 ID>` 可以按注册表顺序运行子集规则，`--severity-threshold <info|warning|error>` 可以过滤返回的 findings。`data.summary` 只统计最终返回的 findings，而 `data.stats.finding_count_before_threshold` 会保留阈值过滤前的总数，`data.stats.rules_run` 则记录规范化后的规则顺序。

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

无界面成功响应遵循稳定的信封结构：

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

`inspect columns` 分析工作表中的每一列，帮助你为后续命令选择稳定的字段名。响应数据包含 `columns`，每列包含 `index`、原始 `name`、生成的 `safe_name`、`is_duplicate`、尽力推断的 `inferred_type`、`non_null_ratio`、`formula_ratio` 和 `sample_values`。响应元数据包含 `header_row_mode`、`resolved_header_row`、`column_count` 和 `data_row_count`。

```bash
excel-cli inspect columns path/to/your/file.xlsx --sheet Orders --header-row auto
excel-cli inspect columns path/to/your/file.xlsx --sheet Orders --header-row 2 --format text
```

`inspect tables` 检测工作表中的连续表格区域。响应数据包含 `data.candidates`；每个候选区域包含 `range`、`header_row`、`column_count`、`row_count` 和 `confidence`。响应元数据包含 `candidate_count`。

```bash
excel-cli inspect tables path/to/your/file.xlsx --sheet Orders
excel-cli inspect tables path/to/your/file.xlsx --sheet Orders --format text
```

无界面错误响应遵循稳定的信封结构：

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

应用具有简洁直观的界面：

- **标题栏与工作表标签**：显示当前文件名和所有可用工作表，当前工作表高亮显示
- **电子表格区域**：显示 Excel 数据的主区域
- **内容面板**：显示当前选中单元格的完整内容
- **通知面板**：显示操作反馈和系统通知
- **状态栏**：显示操作提示和当前输入的命令

## 键盘快捷键

- `h`、`j`、`k`、`l` 或方向键：在单元格之间移动（每次 1 格）
- `[`：切换到上一个工作表（停在第一个工作表）
- `]`：切换到下一个工作表（停在最后一个工作表）
- `0`：跳转到当前行的第一列
- `^`：跳转到当前行的第一个非空列
- `$`：跳转到当前行的最后一列
- `gg`：跳转到当前列的第一行
- `G`：跳转到当前列的最后一行
- `Ctrl+←`（Mac 上为 `Command+←`）：如果当前单元格为空，跳到左侧第一个非空单元格；如果当前单元格非空，跳到左侧最后一个非空单元格
- `Ctrl+→`（Mac 上为 `Command+→`）：如果当前单元格为空，跳到右侧第一个非空单元格；如果当前单元格非空，跳到右侧最后一个非空单元格
- `Ctrl+↑`（Mac 上为 `Command+↑`）：如果当前单元格为空，跳到上方第一个非空单元格；如果当前单元格非空，跳到上方最后一个非空单元格
- `Ctrl+↓`（Mac 上为 `Command+↓`）：如果当前单元格为空，跳到下方第一个非空单元格；如果当前单元格非空，跳到下方最后一个非空单元格
- `Enter`：编辑当前单元格
- `y`：复制当前单元格内容
- `d`：剪切当前单元格内容
- `p`：将剪贴板内容粘贴到当前单元格
- `u`：撤销上一次操作（编辑、行/列变更、工作表创建/删除）
- `Ctrl+r`：重做上一次撤销的操作
- `/`：开始向前搜索
- `?`：开始向后搜索
- `n`：跳转到下一个搜索结果
- `N`：跳转到上一个搜索结果
- `:`：进入命令模式（类 Vim 命令）

## Vim 编辑模式

编辑单元格内容时（按 `Enter` 进入编辑模式）：

- **模式切换**：

  - `Esc`：退出 Vim 模式并保存更改
  - `i`：进入插入模式
  - `v`：进入可视模式

- **导航（普通模式下）**：

  - `h`、`j`、`k`、`l`：左、下、上、右移动光标
  - `w`：移动到下一个单词
  - `b`：移动到单词开头
  - `e`：移动到单词末尾
  - `$`：移动到行尾
  - `^`：移动到行首第一个非空白字符
  - `gg`：移动到第一行
  - `G`：移动到最后一行

- **编辑操作**：

  - `x`：删除光标下的字符
  - `D`：删除到行尾
  - `C`：修改到行尾
  - `o`：在下方打开新行并进入插入模式
  - `O`：在上方打开新行并进入插入模式
  - `A`：在行尾追加
  - `I`：在行首插入

- **可视模式操作**：

  - `y`：复制（yank）选中的文本
  - `d`：删除选中的文本
  - `c`：修改选中的文本（删除并进入插入模式）

- **操作符命令**：

  - `y{motion}`：复制 motion 指定的文本
  - `d{motion}`：删除 motion 指定的文本
  - `c{motion}`：修改 motion 指定的文本

- **剪贴板操作**：

  - `p`：粘贴复制或删除的文本
  - `u`：撤销上一次更改
  - `Ctrl+r`：重做上一次撤销的更改

## 搜索模式

按 `/`（向前搜索）或 `?`（向后搜索）进入搜索模式：

- 输入搜索关键词
- `Enter`：执行搜索并跳转到第一个匹配项
- `Esc`：取消搜索
- `n`：跳转到下一个匹配项（搜索执行后）
- `N`：跳转到上一个匹配项（搜索执行后）
- 搜索结果以黄色高亮显示
- 搜索采用先行后列的顺序（从左到右逐行搜索，然后移动到下一行）

## 命令模式

按 `:` 进入命令模式。可用命令：

### 列宽命令

- `:cw fit` - 自动调整当前列宽以适应内容
- `:cw fit all` - 自动调整所有列宽以适应内容
- `:cw min` - 最小化当前列宽（最大 15 或内容宽度）
- `:cw min all` - 最小化所有列宽（最大 15 或内容宽度）
- `:cw [数字]` - 将当前列宽设置为指定值

### JSON 导出命令

- `:ej [h|v] [行数]` - 将当前工作表数据导出为 JSON 格式

  - `h|v` - 表头方向：`h` 为横向（顶部行），`v` 为纵向（左侧列）
  - `行数` - 表头行数（横向）或列数（纵向）

- `:eja [h|v] [行数]` - 将所有工作表导出到单个 JSON 文件
  - 使用与 `:ej` 相同的参数
  - 创建一个 JSON 对象，以工作表名称为键，工作表数据为值

输出文件名按以下格式自动生成：

- 单个工作表：`原文件名_sheet_工作表名称_YYYYMMDD_HHMMSS.json`
- 所有工作表：`原文件名_all_sheets_YYYYMMDD_HHMMSS.json`

JSON 文件保存在与原始 Excel 文件相同的目录中。

### 类 Vim 命令

- `:w` - 保存文件但不退出
- `:wq` 或 `:x` - 保存并退出
- `:q` - 退出（如果有未保存的更改会警告）
- `:q!` - 强制退出不保存
  文件保存逻辑详见[下文](#文件保存逻辑)。

- `:y` - 复制当前单元格内容
- `:d` - 剪切当前单元格内容
- `:put` 或 `:pu` - 将剪贴板内容粘贴到当前单元格
- `:[单元格]` - 跳转到指定单元格（例如 `:A1`、`:B10`）。支持大小写字母（`:a1` 与 `:A1` 效果相同）

### 工作表管理命令

- `:addsheet [名称]` - 在当前工作表后添加新工作表
- `:sheet [名称/编号]` - 按名称或索引切换工作表（从 1 开始计数）
- `:delsheet` - 删除当前工作表

### 行列管理命令

- `:dr` - 删除当前行
- `:dr [行号]` - 删除指定行（例如 `:dr 5` 删除第 5 行）
- `:dr [起始] [结束]` - 删除行范围（例如 `:dr 5 10` 删除第 5 到 10 行）
- `:dc` - 删除当前列
- `:dc [列]` - 删除指定列（例如 `:dc A`、`:dc a` 或 `:dc 1` 均删除 A 列）
- `:dc [起始] [结束]` - 删除列范围（例如 `:dc A C` 或 `:dc a c` 删除 A 到 C 列）

### 其他命令

- `:nohlsearch` 或 `:noh` - 关闭搜索高亮
- `:help` - 显示可用命令
- `:preview` 或 `:pv` - 显示当前工作表目标和样本行的只读预览

## 文件保存逻辑

Excel-CLI 采用非破坏性方式保存文件：

- 保存文件时（使用 `:w`、`:wq` 或 `:x`），应用会检查是否进行了更改
- 如果没有更改，不会创建新文件，并显示"No changes to save"消息
- 如果启用了懒加载，所有未加载的工作表会在保存前加载，以保留工作簿内容
- 如果进行了更改，会创建一个带时间戳的新文件，格式为 `原文件名_YYYYMMDD_HHMMSS.xlsx`
- 新文件不包含任何样式
- 原始文件永远不会被修改

## 贡献指南

分支命名、提交信息和 Pull Request 规范请参阅 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 许可证

MIT
