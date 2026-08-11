# zotero-literature-docs

一个面向本地 Zotero 文献库的 Codex Skill：**检索 → 分析 → 综述 → 一键产出 Word + PDF + PPT 三件套**。

从 Zotero 本地 API 读取干净数据，生成符合中、英文期刊排版规范的 Word、PDF 和 PPT 文件，并从文献 PDF 中裁剪**真实插图**，装配成带「图片引自参考文献 [n]」标注的汇报 PPT。所有产物自动校验（重开校验 + 乱码扫描 + PowerPoint 打开并导出 PDF 预览）。

> [English](README.en.md) | 中文

## 功能特性

- **本地检索**：`search` 按关键词列出 Zotero 库中匹配条目（key / 年份 / 标题）
- **干净数据**：从本地 API 拉取 JSON（`fetch`），自动清洗脏标题、跳过失效条目；不使用 `export-bibtex` 文本（特殊字符易乱码）
- **正文速读**：`texts` 导出每篇 PDF 的前两页 + 末页文本，写综述前快速读摘要 / 引言 / 结论
- **综述生成**：内容写在 `work/review_content.json`，正文用 `[@ZoteroKey]` 标注引用 → 一键生成 Word + PDF
  - 作者-年份文中引用（Zhang et al.，2015），连续引用合并，参考文献表按作者字母排序
  - 引用了但数据缺失会报错；拉取了但未引用的条目会警告，避免"幽灵引用"
  - 中文排版：宋体 / 黑体、编号标题、首行缩进、三线表；基因名与拉丁学名斜体、蛋白质正体
- **真实插图**：`figs` 按 PDF 内嵌图片的包围盒裁剪真实图（非整页截图）→ `work/ppt_images/`
- **汇报 PPT**：`ppt` 基于 PowerPoint 生成骨架，JSON 配置驱动（cover / toc / pic / text / refs 五种版式），支持任意页数
- **质量校验**：Word / PDF 重开校验 + 乱码扫描；PPT 以有窗口方式在 PowerPoint 中打开并导出 PDF 预览
- **中英双语**：配置 `"lang": "en"` 即可输出英文版三件套（Abstract / Keywords / References 标签与 PPT 图注自动切换为英文）

## 环境要求

- Windows（已验证）｜Python 3.10+（开发环境 3.14）
- Zotero Desktop 已启动，本地 API 端口 23119
- 字体：宋体（simsun.ttc）、黑体（simhei.ttf），PDF 可选等线（Deng.ttf / Dengb.ttf）；Word / PPT 用微软雅黑
- PowerPoint（仅用于 PPT 校验，不参与文件转换）
- 依赖：fpdf2、fontTools、PyMuPDF/fitz、python-docx、lxml（Skill 会在项目 `work/` 下放置可用的离线依赖目录）

## 快速开始

在项目根目录运行（首次会自动把流水线脚本与骨架复制到 `work/`）：

```bash
# 1. 一次性初始化
python <skill>/scripts/zotero_docs.py setup

# 2. 检索 Zotero 本地库（只读）
python work/zotero_docs.py search "anthocyanin"

# 3. 拉取干净数据（key 用逗号分隔）
python work/zotero_docs.py fetch KEY1,KEY2,...

# 4. （可选）导出 PDF 头尾页文本，先读再写
python work/zotero_docs.py texts KEY1,KEY2,...

# 5. 编写 work/review_content.json（分节 + [@KEY] 引用）后生成综述
python work/zotero_docs.py review work/review_content.json

# 6. 裁剪论文真实插图
python work/zotero_docs.py figs KEY1,KEY2,...

# 7. 编写 work/deck_config.json 后生成 PPT
python work/zotero_docs.py ppt work/deck_config.json
```

输出物在 `outputs/`：`<名称>.docx`、`<名称>.pdf`、`<名称>.pptx`。

## 目录结构

```text
zotero-literature-docs/
├── SKILL.md                        # Skill 定义（Codex 读取入口）
├── agents/openai.yaml              # 触发与行为配置
├── references/method.md            # 完整方法、依赖与踩坑记录（中文）
├── references/method.en.md         # 完整方法、依赖与踩坑记录（英文版）
├── scripts/                        # 流水线脚本（统一 CLI：zotero_docs.py）
│   ├── zotero_docs.py              # setup/status/search/fetch/texts/review/figs/ppt
│   ├── fetch_items.py              # 数据拉取与清洗
│   ├── make_review_docs.py         # 综述 Word + PDF 生成
│   ├── extract_figures.py          # PDF 真实插图裁剪
│   ├── build_pptx.py / assemble_pptx.py  # PPT 生成与装配
│   └── deck_config.example.json    # PPT 配置示例
└── assets/_skeleton.pptx           # PowerPoint 生成的骨架模板
```

## 安装为 Codex Skill

- 方式一：将本仓库的 `zotero-literature-docs` 目录复制到 `~/.codex/skills/`，重启 Codex
- 方式二：通过 Codex 技能安装器从本仓库安装
- 使用：新建会话直接描述任务即可自动触发，例如「把Zotero 里关于***(关键词)的文献整理成综述，并生成 Word / PDF / PPT」

## 注意事项

- 数据源一律使用 Zotero 本地 API JSON，不要用 `export-bibtex` 文本
- 新主题出现新的基因名 / 植物属名时，在 `scripts/make_review_docs.py` 的 `ITALIC_WORDS` / `PLANT_GENERA` 中补充
- 不要用 Word COM 做文档转换（本机不稳定）；PPT 校验必须 `WithWindow=$true`
- 配图默认裁剪真实插图；仅当论文无内嵌位图且用户明确同意时才使用 `--whole-page` 整页模式
- 需要英文版交付物时，在 `review_content.json` 与 `deck_config.json` 中设置 `"lang": "en"`

## 开源协议

MIT License，见 [LICENSE](LICENSE)。
