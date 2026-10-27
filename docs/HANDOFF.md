# Agent 工作交接

本文件只记录当前工作状态，不保存长期项目规范。长期规范见仓库根目录 `AGENTS.md` 与 `docs/specs/`。新 Agent 必须用当前代码、`git status` 和相关差异核对以下内容。

## 当前目标与验收条件

- 把 `AGENTS.md` 里的五块功能细则拆到 `docs/specs/`，让每次请求常驻的只有强制规则和规格索引。
- 验收条件：`AGENTS.md` 只剩强制规则、读取顺序和规格索引；五份规格逐条承接原章节；三个工具入口改为按索引读取。

## 已完成

- 新建 `docs/specs/ui-common.md`、`sql-query-form.md`、`batch-refresh.md`、`preview-and-write.md`、`ai-assistant.md`，规则文字按原顺序逐条搬入，没有重写。
- 连接树注释显示那一条从原「预览与写入」移到 `ui-common.md`，因为 SQL 查询和批量刷新共用这棵树，放在写入路径下会在改连接树时被漏读。
- `AGENTS.md` 从 35,976 字节降到 7,372 字节，新增「读取顺序」「规格索引」；删掉界面、SQL 查询窗体、批量刷新、预览与写入、AI 助手五节。
- 读取顺序里写明：动手前按索引读完命中规格，命中多份都读；最终回复用一行列出已读规格；`AGENTS.md` 超过约 15 KB 就把细则移回 `docs/specs/`。
- 把「唯一的 AI 项目约定正文」和「细则只写在本文件」两处旧口径改成「强制规则只在 `AGENTS.md`，功能细则只在 `docs/specs/`，一条规则只写一处」。
- 更新 `CLAUDE.md`、`.cursor/rules/db2sheet.mdc`、`.github/copilot-instructions.md` 的入口说明，保留 `CLAUDE.md` 的 `@AGENTS.md` 导入。
- 修正 `docs/UI-UNIFICATION-PLAN.md` 的界面细则指针和本文件开头的长期规范指针。
- 保留上一轮未提交的命令注释约定（`<# 用途｜风险 #>`），本轮一并入库。

## 本次涉及文件

- `AGENTS.md`
- `docs/specs/ui-common.md`（新增）
- `docs/specs/sql-query-form.md`（新增）
- `docs/specs/batch-refresh.md`（新增）
- `docs/specs/preview-and-write.md`（新增）
- `docs/specs/ai-assistant.md`（新增）
- `CLAUDE.md`、`.cursor/rules/db2sheet.mdc`、`.github/copilot-instructions.md`
- `docs/UI-UNIFICATION-PLAN.md`、`docs/HANDOFF.md`

## 验证结果

- 用 HEAD 版 `AGENTS.md` 按 UTF-8 逐行比对：127 行里除 8 行按计划改写（开头两段、操作说明、批量刷新与预览写入各一处跨节引用、文档章节三条），其余每条规则都能在新文件里逐字找到。
- 同一脚本检查原文规则是否在多处重复出现，重复 0 条。
- 规格索引、五个文件头的触发路径，以及正文里引用的源码路径都用 `Test-Path` 核对过，全部存在。
- 未运行 `MSBuild` 和发布脚本：纯文档改动，不涉及编译产物。
- 未做子 Agent 冷启动试验，需要新会话；留待下次按「改批量刷新状态圆点颜色」验证索引是否够清楚。

## 阻塞项与下一步

- 无阻塞。
- `ai-assistant.md` 是最大的一份，仍承接原 AI 章节全部内容；若以后只改请求或设置部分，可再按「对话栏外观」与「模型、Skill、请求与审计」拆开。
- 下次改 UI 或 AI 相关文件时按索引读规格，并在回复里列出已读规格；连续几次出现漏读再调整索引写法。
- 下次发布仍先读 `.cursor/skills/publish-release/SKILL.md`；只改 `Properties/AssemblyInfo.cs` 两处版本号，先跑 `scripts/publish-release.ps1 -CheckOnly`，再发布并核验。
