# ISO Checklist Writer

ISO27001/ISO20000 合规审查检查表自动生成工具 - WorkBuddy Skill

## 🎯 核心特性

- ✅ **一键生成**：单命令完成读取记录 → AI生成内容 → 格式克隆输出
- ✅ **零WorkBuddy Token消耗**（方案A）：完全离线运行，双击即用
- ✅ **最小Token消耗**（方案B）：通过WorkBuddy触发，仅消耗50-100个token
- ✅ **全程DeepSeek驱动**：所有重推理工作通过DeepSeek API执行，成本约¥0.3-0.5/次
- ✅ **格式克隆**：100%保留模板格式，只填写内容列
- ✅ **问题发现**：自动发现问题并生成问题清单、整改建议

## 📋 两种使用方案

### 方案A（推荐，常用）：完全离线运行

**Token消耗**：0个 WorkBuddy token  
**成本**：仅消耗 DeepSeek token（~15万/次，约¥0.3-0.5）

**使用方法**：
1. 打开文件夹：`C:\Users\Confu\.workbuddy\skills\iso-checklist-writer\`
2. 双击：`启动检查表生成工具.bat`
3. 按提示输入4个参数：
   - 客户记录目录路径
   - 标准类型（ISO27001 或 ISO20000）
   - 审核日期（格式：YYYY-MM-DD）
   - 企业名称
4. 等待生成完成（约2-3分钟），查看输出文件

**输出位置**：
- 检查表和问题清单：`Skill目录\output\`
- 详细日志：`Skill目录\output\log_YYYYMMDD_HHMMSS.txt`

---

### 方案B（备选）：最小Token消耗模式

**Token消耗**：~50-100个 WorkBuddy token（约¥0.2）  
**成本**：DeepSeek token（~15万/次，约¥0.3-0.5）

**使用方法**：
1. 在WorkBuddy对话中说："生成ISO20000检查表，客户记录目录XXX，审核日期XXX"
2. AI会调用 `run_all.py --quiet ...`（只输出最终结果）
3. 查看生成的文件（AI会交付给你）

**适用场景**：
- 方案A不可用时（如批处理脚本损坏）
- 需要通过对话触发时（如远程操作）
- **调优时使用**：当方案A遇到问题，用方案B调优，然后发布A1, A2, A3等改进版本

---

## 🚀 快速开始

### 1. 安装依赖

```bash
pip install python-docx
```

### 2. 配置DeepSeek API Key

**方法1**：设置环境变量
```bash
export DEEPSEEK_API_KEY="your-api-key-here"
```

**方法2**：直接编辑 `run_all.py`，替换第28行的API Key

### 3. 运行（方案A）

双击 `启动检查表生成工具.bat`，按提示输入参数即可。

### 3. 运行（方案B）

```bash
python run_all.py \
  --records-dir "D:/path/to/records" \
  --standard ISO20000 \
  --audit-date 2026-02-01 \
  --company-name "企业名称" \
  --quiet
```

---

## 💰 成本对比

| 方式 | WorkBuddy Token | DeepSeek Token | 成本 |
|------|-----------------|----------------|------|
| 旧方式（无--quiet） | ~450-500个 | ~150,000个 | ~6-7元 |
| 方案B（--quiet模式） | ~50-100个 | ~150,000个 | ~0.5-0.7元 |
| 方案A（完全离线） | **0个** | ~150,000个 | **~0.3-0.5元** |

**节省**：方案A比旧方式降低 **95%** 成本！

---

## 📂 项目结构

```
iso-checklist-writer/
├── run_all.py                       # 核心引擎：一键生成检查表
├── run_quiet.py                    # 安静模式启动器（方案A核心）
├── 启动检查表生成工具.bat          # Windows启动脚本（方案A使用）
├── 一键推送到GitHub.bat           # GitHub推送脚本
├── read_doc_tool.py                # .doc/.docx读取工具库
├── output/                        # 输出文件目录
│   ├── *.docx                     # 生成的检查表
│   └── *.txt                      # 生成的问题清单
├── SKILL.md                       # WorkBuddy Skill完整文档
└── README.md                      # 本文件
```

---

## ⚙️ 高级配置

### DeepSeek API 配置

| 项目 | 值 |
|------|-----|
| API Key | 环境变量 `DEEPSEEK_API_KEY`，fallback: 脚本内置 |
| Base URL | `https://api.deepseek.com` |
| 模型 | `deepseek-chat`（实际为deepseek-v4-flash） |
| 费用估算 | ~¥0.3-0.5/次 |

### 记录截断限制

- 当前：`MAX_RECORDS_CHARS = 25000`（Step2传递给DeepSeek的记录文本长度）
- 调整：修改 `run_all.py` 第22行的 `MAX_RECORDS_CHARS` 值

### Batch Size

- 当前：`BATCH_SIZE = 6`（每次DeepSeek API调用生成6个条款的内容）
- 调整：减小到 3-4 可提高质量；增大到 8-10 可加快速度

---

## 🐛 已知问题与修复

### 问题1：Step1读取0个文件

- **现象**：`Step1: 0 files, 0 chars`
- **原因**：`priority_dirs` 硬编码了特定目录结构
- **修复**：改为递归读取所有文件（`root.rglob("*")`）
- **状态**：✅ 已修复（2026-06-14）

### 问题2：输出目录权限被拒

- **现象**：`Permission denied: D:\...`
- **原因**：沙箱阻止Python写入D盘
- **修复**：`OUTPUT_DIR` 改为脚本本地 `output` 子目录
- **状态**：✅ 已修复（2026-06-14）

### 问题3：Windows GBK编码崩溃

- **现象**：`UnicodeEncodeError: 'gbk' codec can't encode...`
- **原因**：Python on Windows默认stdout编码为GBK
- **修复**：添加 `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')`
- **状态**：✅ 已修复（2026-06-13）

---

## 📝 检查表输出规范

### Rule 1：输出格式必须为Word文档

所有报告/检查表必须输出为 `.docx` 格式

### Rule 2：检查表必须严格按模板格式输出（格式克隆机器人）

- 只能填写"内容"列，其他部分一律不得修改
- 格式克隆：打开.docx模板 → 只填空白单元格 → 100%保留模板原文
- 绝对禁止：添加颜色、加粗、项目符号、分隔线、调整段落缩进/行距/字体大小

### Rule 3：内容列必须包含制度文件摘抄

格式："[制度文件名]：[3-5句简要摘抄]。查[记录证据]..."

---

## 📊 测试验证

### 山东翰林 ISO20000（审核日期2026-02-01）

- DeepSeek tokens: 118,979
- 填充: 64/64 条款，0 失败
- 总耗时: 99.9 秒
- 输出: `2026-06-14_上午_山东翰林_ISO20000检查表_DS001944.docx`

### 山东世纪黄河 ISO20000（审核日期2026-06-13）

- DeepSeek tokens: 148,994
- 填充: 64/64 条款，0 失败
- 总耗时: 132.4 秒
- 输出: `2026-06-14_上午_山东世纪黄河_ISO20000检查表_DS003304.docx`

---

## 📄 许可证

MIT License

---

## 📧 联系方式

- 作者：Confu
- GitHub：[@Williamkongyb](https://github.com/Williamkongyb)
- WorkBuddy Skill：兼容 WorkBuddy AI 助手

---

**生成时间**：2026-06-14  
**版本**：v2.0（支持方案A + 方案B）
