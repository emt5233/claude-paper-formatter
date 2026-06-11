---
name: paper-formatter
description: >
  Formats user-provided content into a Chinese academic paper .docx document
  according to strict formatting specifications. This skill should be used when
  users ask to format a paper, generate an academic article, create a formatted
  thesis/manuscript, or produce a .docx document following Chinese academic
  journal standards. Triggers on mentions of "论文格式", "格式化论文", "生成论文",
  "学术论文排版", "论文docx", "整理成论文", or similar requests involving
  Chinese academic paper formatting.
---

# Paper Formatter

Format user-provided content into a strictly-formatted Chinese academic paper
.docx document according to the specification in `references/formatting_spec.md`.

## Workflow

### Step 1: Gather and Structure Content

Query the user for the following required elements. If the user has already
provided raw content, organize it into this structure:

- **Title (Chinese)**: ≤20 characters, no uncommon English abbreviations
- **Title (English)**: Matching the Chinese title
- **Chinese Abstract**: ~200 characters, third-person, include innovation points
- **English Abstract**: Matching the Chinese abstract, innovations in italics
- **Chinese Keywords**: 3-8 keywords, semicolon-separated
- **English Keywords**: Matching Chinese keywords, lowercase unless acronyms
- **Body Sections**: At most 3 heading levels (no level 4+)
- **References**: In the standard format (see formatting spec)
- **Optional**: Tables, figures, equations, appendices, acknowledgements

### Step 2: Validate Content

Check the following before generating:

1. Title does not exceed 20 Chinese characters
2. Abstract is ~200 characters and written in third person (no 本文/论文)
3. Keywords are 3-8 items
4. Section headings use at most 3 levels
5. All references have complete metadata

### Step 3: Build JSON Input

Organize all content into the JSON structure expected by
`scripts/generate_paper.py`. See `references/formatting_spec.md` for the
detailed field descriptions and formatting rules.

The JSON schema:

```json
{
  "title_cn": "不超过20字的中文题目",
  "title_en": "English Title Matching Chinese",
  "abstract_cn": {
    "text": "摘要正文约200字...",
    "innovations": ["创新点1", "创新点2"]
  },
  "abstract_en": {
    "text": "English abstract matching Chinese...",
    "innovations": ["Innovation point 1", "Innovation point 2"]
  },
  "keywords_cn": ["关键词1", "关键词2", "关键词3"],
  "keywords_en": ["keyword1", "keyword2", "keyword3"],
  "sections": [
    {
      "level": 1,
      "title": "一级标题",
      "children": [
        {
          "type": "paragraph",
          "text": "正文段落内容，参考文献用[1]上标标注。"
        },
        {
          "type": "table",
          "caption": "表1.1 表的名称",
          "headers": ["列1", "列2", "列3"],
          "rows": [["值1", "值2", "值3"], ["值4", "值5", "值6"]],
          "notes": ["注释①", "注释②"]
        },
        {
          "type": "figure",
          "caption": "图1 图的名称",
          "description": "图片描述或替代文字"
        },
        {
          "type": "equation",
          "latex": "E = mc^2"
        },
        {
          "type": "section",
          "level": 2,
          "title": "二级标题",
          "children": [...]
        }
      ]
    }
  ],
  "references": [
    {
      "type": "journal",
      "authors": "作者1, 作者2",
      "title": "文献题名",
      "journal": "刊名",
      "year": "2024",
      "volume": "42",
      "issue": "3",
      "pages": "100-110"
    },
    {
      "type": "book",
      "authors": "作者",
      "title": "书名",
      "publisher": "出版社",
      "location": "出版地",
      "year": "2023",
      "pages": "50-60"
    }
  ],
  "appendices": [
    {"title": "附录1 标题", "content": ["段落1", "段落2"]}
  ],
  "acknowledgement": "致谢内容"
}
```

Reference types: `journal`, `book`, `conference`, `dissertation`, `newspaper`,
`electronic`, `patent`, `standard`, `other`.

### Step 4: Generate the .docx

Run the generation script:

```bash
python scripts/generate_paper.py input.json output.docx
```

The script applies all formatting from the specification:
- Page margins: top/bottom 2.54cm, left/right 3.17cm
- Title: 三号 黑体 加粗 居中
- Abstract headings: 黑体 小四 加粗 左对齐
- Abstract body: 宋体 小四
- Keywords: 黑体/宋体 小四
- Section headings: 黑体/宋体 with proper sizes
- Body text: 宋体 小四, 1.5x line spacing
- References: 五号 宋体
- Tables: 三线表 format
- Page numbers: Roman for front matter, Arabic for body

### Step 5: Verify Output

After generation, check:
1. The .docx opens correctly
2. Font sizes, faces, and styles match the specification
3. Page margins and numbering are correct
4. All sections are present and properly ordered

## Bundled Resources

- `scripts/generate_paper.py`: Python script that generates the .docx from a
  JSON input file. Requires `python-docx`. Install with `pip install python-docx`.
- `references/formatting_spec.md`: Complete formatting specification with all
  font sizes, faces, spacing, margins, and reference format details.
