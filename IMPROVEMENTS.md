# Skill 改进记录（zotero-literature-docs）

> 本文件记录 Skill 的改进方法，与仓库内容同步；发布时如需公开可一并提交或删除。

## 2026-08-10

1. **英文模式（`"lang": "en"`）**
   - `make_review_docs.py`：摘要 / 关键词 / 参考文献标签切换为 Abstract / Keywords / References；页脚切换为 Page X of Y；PDF 正文、标题、页脚使用 Times New Roman（TimesR/TimesB）；图中英文题注使用 Times。
   - `build_pptx.py`：目录、参考文献标题与图注切换为英文（"Figure from reference [n]"）。
   - 配置方法：`review_content.json` 与 `deck_config.json` 顶部加 `"lang": "en"`。
2. **中英文参考文献著录统一英文标点**
   - `ref_text` 统一使用英文标点：作者间英文逗号、卷期紧贴（如 51(03)）、页码与 DOI 英文冒号；中英文模式一致。
   - 文中引用保持各自语言格式：中文（作者，年份），英文 (Author, year)。
3. **中英双语文档**
   - 新增 README.en.md、PUBLISHING.en.md、references/method.en.md；README 互链切换。
4. **实测记录**
   - 测试主题1（3 篇）：中英文三件套，PPT 含 5 张真实插图。
   - 测试主题2（13 篇）：中文综述 Word/PDF/MD（正文约 4000 字，62 处引注），验证通过。
5. **待推送提交**：README/PUBLISHING/method.en/英文模式/标点修正等本地提交，网络恢复后推送。
