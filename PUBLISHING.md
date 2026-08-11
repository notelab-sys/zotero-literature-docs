# 发布说明（Publishing Notes）

当前状态：仓库为 **Private**，本文档为发布准备材料。确定发布后按下方清单执行。

## 版本计划

### v0.1.0（拟定初版）

- 功能定位：本地 Zotero 数据库检索 → 分析 → 综述 → 生成 **Word + PDF + PPT 三件套**，便于使用
- 核心差异：一套流水线产出三件套；PPT 使用从文献 PDF 裁剪的真实插图，并标注「图片引自参考文献 [n]」；所有产物自动校验
- 目标用户：中文科研用户（植物学 / 园艺学方向），输出符合中文期刊排版规范
- 兼容性：Windows + Codex（桌面 / CLI）；Python 3.10+；Zotero 本地 API（端口 23119）

## 发布前检查清单

1. **可见性**：仓库 Settings → 改为 Public；发布前再确认无敏感信息（已扫描，仅含系统字体路径）
2. **许可证**：确认 MIT（见 `LICENSE`，版权人 konjac2027）；如需其他协议请先替换
3. **README 复核**：功能、快速开始、环境要求、目录结构与实际内容一致
4. **脚本自检**：在干净环境跑通 `zotero_docs.py status`，确认依赖说明完整（fpdf2 / fontTools / PyMuPDF / python-docx / lxml）
5. **版本标记**：`git tag v0.1.0` 并推送；可选创建 GitHub Release（附三件套示例输出截图或样例文件）
6. **可选**：接入 Codex 技能市场 / 插件市场（marketplace manifest），支持一键安装
7. **登记**：发布后在工作区全局日志 / 全局控制台登记发布状态与仓库链接

## 用户如何安装

- 方式一（推荐）：Codex 技能安装器，输入本仓库路径一键安装
- 方式二：手动复制 `zotero-literature-docs` 目录到 `~/.codex/skills/`，重启 Codex
- 使用：新建会话直接描述任务即可触发，例如「把 Zotero 里关于花青素的文献整理成综述并出 Word / PDF / PPT」

## 维护约定

- Skill 内容（SKILL.md / scripts / references）每次修改后同步提交到本仓库
- 重要行为变化：更新 README 并递增版本号；踩坑记录追加到 `references/method.md`
- 三件套生成逻辑的调整优先沉淀为可复用脚本，避免每次手工处理
