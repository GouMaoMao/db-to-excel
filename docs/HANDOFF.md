# Agent 工作交接

本文件只记录当前工作状态，不保存长期项目规范。长期规范见仓库根目录 `AGENTS.md`。新 Agent 必须用当前代码、`git status` 和相关差异核对以下内容。

## 当前目标与验收条件

- 当前没有进行中的编码任务。

## 已完成

- 将跨 Agent 项目约定统一到根目录 `AGENTS.md`。
- 将 Copilot 与 Claude Code 的专用文件改为公共约定入口。
- 保留 `.cursor/skills/publish-release/SKILL.md` 作为所有 Agent 发布时必须读取的专项流程。

## 待完成

- 无。

## 本次涉及文件

- `AGENTS.md`
- `.github/copilot-instructions.md`
- `CLAUDE.md`
- `docs/HANDOFF.md`

## 验证结果

- 已核对公共约定、工具入口和交接文件的路径及内容。
- 本次只修改 Markdown 文档，不需要构建程序。

## 阻塞项与下一步

- 无阻塞项。
- 开始下一项任务时，用实际任务目标、验收条件和状态替换本文件对应内容；任务结束前再次更新。