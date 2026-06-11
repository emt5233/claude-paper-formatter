# Claude Paper Formatter

Claude Code 技能：将用户提供的内容格式化为**中文学术论文 .docx**，严格遵循学术期刊排版规范。

## 功能特性

- 📄 **一键生成 .docx**：输入 JSON 内容，输出完全格式化的学术论文文档
- 📐 **严格格式规范**：页边距、字体、字号、行距、页码全面符合中文学术期刊标准
- 📊 **三线表支持**：自动生成并格式化学术三线表
- 🔢 **公式排版**：公式居中、编号右对齐
- 📖 **参考文献格式化**：支持期刊[J]、书籍[M]、会议[C]、学位论文[D]、专利[P] 等 9 种类型
- 🏷️ **罗马/阿拉伯页码**：前页罗马数字，正文阿拉伯数字
- 📑 **自动目录**：根据章节层级自动生成带页码的目录

## 格式规范

严格遵循以下排版标准（详见 `references/formatting_spec.md`）：

| 元素 | 字体 | 字号 |
|------|------|------|
| 中文标题 | 黑体 加粗 居中 | 三号 (16pt) |
| 摘要正文 | 宋体 | 小四 (12pt) |
| 一级标题 | 黑体 加粗 | 四号 (14pt) |
| 二级标题 | 黑体 加粗 | 小四 (12pt) |
| 正文 | 宋体 | 小四 (12pt) |
| 参考文献 | 宋体 | 五号 (10.5pt) |

- 页边距：上下 2.54cm，左右 3.17cm
- 行距：1.5 倍
- 页码：前页罗马数字（i, ii, iii...），正文阿拉伯数字（1, 2, 3...）
- 表格：三线表格式

## 安装

### 1. 安装 Claude Code Skill

将仓库克隆到 Claude Code 的 skills 目录：

```bash
# 进入你的 Claude Code 项目目录
cd your-project

# 克隆到 skills 目录
git clone https://github.com/emt5233/claude-paper-formatter.git .claude/skills/paper-formatter
```

### 2. 安装 Python 依赖

```bash
pip install python-docx
```

## 使用方法

在 Claude Code 对话中，直接告诉 Claude：

> "帮我把这篇内容格式化成学术论文 docx"
>
> "生成中文学术论文"
>
> "格式化论文"

Claude 会自动触发此技能，引导你提供以下内容：
- 中英文标题（中文 ≤20 字）
- 中英文摘要（约 200 字，第三人称）
- 中英文关键词（3-8 个）
- 正文章节
- 参考文献
- 附录、致谢（可选）

### 手动使用

也可以直接准备 JSON 文件并运行脚本：

```bash
python scripts/generate_paper.py input.json output.docx
```

参考 `assets/sample_input.json` 查看 JSON 输入格式示例。

## 文件结构

```
paper-formatter/
├── SKILL.md                       # Claude Code 技能定义
├── README.md                      # 本文件
├── references/
│   └── formatting_spec.md         # 完整格式规范
├── scripts/
│   └── generate_paper.py          # 生成 .docx 的 Python 脚本
└── assets/
    └── sample_input.json          # 示例输入文件
```

## 支持的参考文献类型

| 类型 | 代码 | 类型 | 代码 |
|------|------|------|------|
| 期刊文章 | J | 专利 | P |
| 书籍 | M | 标准 | S |
| 会议论文 | C | 报告 | R |
| 学位论文 | D | 电子资源 | EB/OL |
| 报纸 | N | 其他 | Z |

## 依赖

- Python 3.7+
- [python-docx](https://python-docx.readthedocs.io/)

## License

MIT
