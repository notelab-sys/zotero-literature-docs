# 从 Zotero 生成 Word / PDF 文献文档的方法（用户指定保留）

## 核心原则

1. **数据必须从 Zotero 本地 API 读取（JSON）**，不要用 `export-bibtex` 导出的文本。
   BibTeX 导出在含特殊符号（如连接号、弯引号、连字 ﬃ 等）时会产生乱码，表现为
   「鈥恊diting」「铿乧iency」这类字符。
2. Word 用 python-docx 生成，PDF 用 fpdf2 生成，两者共用 `work/data.json`。
3. 生成后必须扫描确认没有乱码字符（见下方验证）。

## 步骤

1. 确认 Zotero 本地 API 运行：
   `python3 <zotero-skill>/scripts/zotero.py status --json`
2. 取干净数据（itemKey 用逗号分隔）：
   `python work/fetch_items.py --keys KEY1,KEY2,...` → 生成 `work/data.json`
3. 一键生成并验证（Word + PDF + 乱码扫描）：
   `python work/run_all.py [doc_config_xxx.json]`
   每个主题一份配置文件（标题 + 分节 + 条目 key），未指定参数时使用默认配置。
4. 输出文件在 `outputs/`：`基因遗传转化文献.docx`、`基因遗传转化文献.pdf`。
5. 如需修复 BibTeX 文件中的乱码字段：`python work/fix_bib.py`。
6. 综述文档：内容写在 `work/review_content.json`，正文用 `[@ZoteroKey]` 标注引用
   （自动编号），运行 `python work/make_review_docs.py` 一键生成综述 Word + PDF，
   并自动附带按出现顺序编号的参考文献列表。
7. PPT：`python work/assemble_pptx.py`（内部调用 `build_pptx.py` 生成幻灯片内容，
   填入 PowerPoint 生成的 `work/_skeleton.pptx` 骨架，保证结构合法）。
   配图用 `work/extract_figures.py` 从文献 PDF 渲染图页，幻灯片上标注“图片引自参考文献 [n]”。

## PPT 踩坑记录（2026-08-04）

- 自建整套 pptx 骨架（theme/master/layout）PowerPoint 不认，务必用 PowerPoint 生成的
  空模板做骨架，再替换 slide XML。
- 文本居中必须是 `algn="ctr"`，写 `algn="c"` 会被判损坏。
- 添加 png 媒体时要在 `[Content_Types].xml` 补 `<Default Extension="png" .../>`。
- 幻灯片关系文件要保留原字节、只追加图片关系，不要整段重写。
- PowerPoint COM 自动化打开含图片的 pptx 必须 `WithWindow=$true`；用无窗口方式会误报
  “无法打开该文件”。COM 的 AddPicture 在本机不稳定、会挂起，不要依赖它。

## 依赖（已装在 work/ 下，无需联网重装）

- `work/pylibs`：fpdf2 + fontTools（PDF 生成）
- `work/pylibs2`：python-docx + lxml（cp314 非 cp314t）+ typing_extensions（Word 生成）
- 若需重装：从 PyPI 下载 wheel 直接解压到对应目录；lxml 必须选
  `cp314-cp314-win_amd64`（不要 cp314t），fpdf2 需要 fontTools。

## 踩坑记录（下次避免）

- PDF 字体用等线 `Deng.ttf` / 粗体 `Dengb.ttf`；不要用 SimHei（fpdf2 multi_cell 会报
  「Not enough horizontal space」）。
- fpdf2 的 `multi_cell` 结束后 x 停在右边界，下一次调用前必须 `set_x(l_margin)`。
- Word 中文字体需在 `rPr/rFonts` 显式设置 eastAsia（Microsoft YaHei）。
- 含中文文件名的脚本请写成 .py 文件再运行（PowerShell 内联 heredoc 管道会按 GBK 损坏中文）。
- 检查输出目录时用 glob 会匹配到 Word 遗留的 `~$` 开头临时锁定文件（不是 zip，
  会被误判为损坏）；先排除或清理。
- 本机 Word COM 自动化（Documents.Open/SaveAs2）不稳定、会挂起，不要依赖它转换文档。

## 更新（2026-08-04）

用户指定（已保存）：以后文献文档的 **Word 和 PDF 一律按本方法生成**——
数据从 Zotero 本地 API 读取 `work/data.json`，Word 用 python-docx、PDF 用 fpdf2，
生成后统一做重开校验与乱码扫描。新任务默认直接走本流程，不再使用其他方式。

**PPT 也按本方法生成**（2026-08-04 追加）：`work/extract_figures.py` 从文献 PDF
提取配图 → `work/assemble_pptx.py`（含 `build_pptx.py`）基于 PowerPoint 生成的
骨架生成 PPT，图片标注“引自参考文献 [n]”，生成后用 PowerPoint 有窗口方式打开验证
并导出 PDF 预览确认。

## 更新（2026-08-05）：体验改进

根据一次完整实战（花青素综述 + MBW 专题 + 两套 PPT）的复盘，做了以下改进：

1. **统一命令行入口**：新增 `scripts/zotero_docs.py`，支持
   `setup / status / search / fetch / texts / review / figs / ppt` 子命令，
   在项目根目录运行即可；首次运行自动把流水线脚本与骨架复制到 `work/`。
2. **配图改为裁剪真实插图**：`extract_figures.py` 默认按 PDF 内嵌图片的
   bounding box 裁剪（`<KEY>_f1.png`），不再整页截图；`--whole-page` 保留为
   显式可选的整页模式。key 通过命令行或 `--keys-file` 传入，不再硬编码。
3. **PPT 内容改为 JSON 配置驱动**：新增 `scripts/deck_config.example.json`，
   幻灯片内容写在 `work/deck_config.json`（cover/toc/pic/text/refs 五种类型），
   无需改 Python 代码。
4. **解除 16 页硬限制**：`assemble_pptx.py` 支持任意页数——超出骨架时克隆
   幻灯片部件并补齐 Content_Types/关系，少于骨架时移除多余部件。
5. **数据获取更健壮**：`fetch_items.py` 对失效/404 条目跳过并警告（不再
   KeyError 崩溃），并清洗标题开头的编号噪音（如 “1 Elucidation of …”）。
6. **引用完整性校验**：`make_review_docs.py` 生成前检查所有 `[@KEY]` 均在
   `data.json` 中（缺失则报错列出），并提示未引用的条目。
7. **阅读环节支持**：新增 `extract_texts.py`，从 PDF 提取前两页与末页文本到
   `work/paper_texts/`，方便撰写综述前阅读摘要、引言与结论。

仍沿用且不可绕过的关键规则：数据必须来自 Zotero 本地 API JSON；Word 中文用
eastAsia=Microsoft YaHei；PDF 用 Deng 字体并每次 multi_cell 后 set_x；PPT 基于
PowerPoint 生成的骨架合并；含图片的 PPT 必须用有窗口方式打开验证。

## 更新（2026-08-05 二次）：PDF 排版与数据健壮性

1. **PDF 排版重做**：综述 PDF 采用新的版式——标题居中加装饰线、引言首行缩进、
   章节标题带左侧色条与细分隔线、条目悬挂缩进、每段首行缩进两字、正文左对齐
   （避免两端对齐产生的行中空白）、参考文献编号悬挂、页脚增加细线与
   “第 n 页 / 共 N 页”。正文行距、边距与留白同步调整。
2. **抓取分批改为 25 个 key/次**：实测 Zotero 本地 API 对过长的 itemKey 列表
   会静默漏掉个别条目（50 个 key 时曾漏掉 1 条），改为小批量请求可避免。

## 更新（2026-08-05 三次）：按中文期刊格式重做综述文档

1. **排版**：PDF 正文采用宋体（SimSun）10.5 pt、行距约 5.6 mm；中文标题黑体
   （SimHei）22 pt；一级标题宋体 14 pt（编号 1、2、…）；二级标题黑体 10.5 pt
   （编号 1.1、1.2、…）；摘要/关键词黑体 9 pt；参考文献宋体 7.5 pt。无页眉，
   页脚仅页码。正文每段首行缩进两字、左对齐。
2. **引用格式**：正文引用改为作者-年份制，如（Zhang et al.，2015）、
   （Porra & Grimme，1978）；参考文献按第一作者姓氏字母排序，格式为
   “作者. 年份. 标题. 期刊，卷(期)：页码. DOI：…”，与中文期刊一致。
3. **图表支持**：综述内容 JSON 的 group 可携带 `table`（caption/caption_en、
   columns、rows、widths）或 `image`（path、width、caption/caption_en）。
   PDF 中表题在表上方、图题在图下方，均为中文+英文对照；表格带网格线，
   表头黑体加粗、表体宋体 7.5 pt。Word 同步渲染表格与图片。
4. 字体依赖：`C:\Windows\Fonts\simsun.ttc`（宋体）与 `simhei.ttf`（黑体），
   缺失时自动回退到 Deng/DengB。

## 更新（2026-08-05 四次）：最终固化（按用户反馈逐项收口）

以下规则均已写入 `make_review_docs.py` 等脚本，新任务默认生效：

1. **引用格式**：正文引用为作者-年份制（（Zhang et al.，2015））；连续多篇引用
   合并为一个括号、以分号分隔；无作者条目显示为“佚名，年份”。参考文献按第一
   作者字母排序，格式“作者. 年份. 标题. 期刊，卷(期)：页码. DOI：…”，最后
   一位作者缩写后加点再接年份；续行悬挂缩进两格。英文标题采用句首大写（句中
   单词小写），基因名保留大写并斜体、拉丁学名属名首字母大写斜体、作者名首字母
   大写正体。
2. **PDF 版式（中文期刊风格）**：宋体正文 10.5 pt、标题黑体 22 pt、一级标题宋体
   14 pt、二级标题黑体 10.5 pt、摘要/关键词黑体 9 pt、参考文献宋体 7.5 pt；
   每段首行缩进两字、续行顶格、左对齐；标题按字符宽度贪婪换行（末尾字符居中
   排第二行，避免短词单独成行）；内容底界 272 mm，杜绝“单行空白页”；
   表格整表同页优先、超一页可跨两页且续页重复表头；表格为三线表；摘要标签与
   内容同行。
3. **Word 版式**：中文字体宋体、英文 Times New Roman；黑体（标题/表头/表题）
   不加粗；页脚“第 X 页 / 共 Y 页”；标题与 PDF 采用相同换行规则；表格行
   cantSplit、除末行外 keepNext，与 PDF 分页行为一致。
4. **斜体规则**：基因名（BBM、LEC1、AP1、LFY 等）与植物拉丁学名（双名法）
   斜体；蛋白质与酶缩写（“蛋白/酶 XX”“XX 蛋白/酶”语境）正体；采用属名白名单
   与常用词过滤避免误标（如 “Current progress”“Genes and” 不斜体）；规则库
   可随新主题扩充。Word 以 run 级斜体实现，PDF 以 Times Italic 渲染斜体拉丁
   片段、宋体渲染中文与正体片段。
5. **PDF 缺字处理**：宋体缺失的拉丁扩展字符（如 ę）经 `pdf_safe` 去音标处理，
   中文标点不受影响。
6. **表格单元格引用**：PDF 与 Word 表格单元格中的 [@KEY] 自动渲染为作者-年份
   引用，可设置“参考文献”列。
7. **健壮性**：所有脚本对 stdout/stderr 做 UTF-8 重配置，避免 GBK 控制台打印
   特殊字符时崩溃；fetch 分批 25 个 key/次；失效条目跳过并警告；标题编号噪音
   自动清洗。
8. **PPT**：幻灯片内容由 deck_config.json 驱动（cover/toc/pic/text/refs），
   任意页数支持；配图默认裁剪真实插图（--whole-page 为显式整页模式）；
   build_pptx 引用统计兼容“仅有表格、无条目”的章节组。

## 更新（2026-08-05 五次）：拉丁斜体与标题规范收口

1. **PDF 拉丁斜体**：PDF 中以 Times Italic（timesi.ttf）渲染基因名与拉丁学名
   等斜体片段，宋体渲染中文与正体片段；Word 以 run 级斜体实现，二者判定规则
   共用 `italic_spans`，保证 Word/PDF 一致（英文摘要亦应用斜体规则）。
2. **参考文献标题**：英文标题句首大写、句中单词小写；全大写缩写/基因名保留
   大写；基因全称（如 AUXIN RESPONSE FACTOR）大写斜体；双名学名属名与种名
   均斜体、命名人（如 Linden.、Pierre ex A. Froehner）保留首字母大写且正体；
   作者名首字母大写正体。
