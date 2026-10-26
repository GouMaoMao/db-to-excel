# Agent 工作交接

本文件只记录当前工作状态，不保存长期项目规范。长期规范见仓库根目录 `AGENTS.md`。新 Agent 必须用当前代码、`git status` 和相关差异核对以下内容。

## 当前目标与验收条件

- 发布 `2.0.2`：提交并推送 GitHub，生成安装包和 GitHub Release，再镜像分支、标签、Release 与附件到 Gitee。
- SQL 查询窗体改用成熟格式化方案，输出紧凑并保留字段或条件后的单行注释。
- 模型设置的检测按钮列出接口返回的全部型号，并提示检测后仍需保存。
- 项目和安装包按 AGPLv3 提供许可证与第三方声明。

## 已完成

- 引入 `PoorMansTSQLFormatter` 1.4.3.1，以 `TSqlStandardFormatter` 替换原约 780 行自写格式化器。
- 固定为两空格缩进、小写关键字、120 字符行宽；关闭字段、CASE、BETWEEN 展开，保留布尔条件分行和行尾注释。
- 「测试」按钮更名为「检测模型」，增加悬浮说明；检测成功后列出接口返回的全部型号。
- 添加 GNU AGPLv3 官方完整 `LICENSE`、`THIRD_PARTY_NOTICES.md` 和 README 许可证入口。
- 版本从 `2.0.1` 升级为 `2.0.2`。
- 本地 Advanced Installer 工程已通过官方 `/AddFile` 命令加入 `PoorMansTSqlFormatterLib35.dll`、`LICENSE` 和 `THIRD_PARTY_NOTICES.md`，均属于 `MainFeature`。
- 更新架构和部署文档。

## 本次涉及文件

- `Contracts/ISqlFormattingService.cs`
- `DB2Sheet.csproj`
- `LICENSE`
- `Properties/AssemblyInfo.cs`
- `README.md`
- `Services/SqlFormattingService.cs`
- `THIRD_PARTY_NOTICES.md`
- `UI/AiModelSettingsForm.cs`
- `UI/AppPresentation.cs`
- `docs/ARCHITECTURE.md`
- `docs/DEPLOYMENT.md`
- `docs/HANDOFF.md`
- `DB2Sheet.Tests/Services/SqlFormattingServiceTests.cs`（测试目录不入库）
- `installer/db-to-excel.aip`、`installer/安装说明.md`（安装目录不入库）

## 验证结果

- NuGet 恢复和主项目 Debug 编译成功，生成新的 `bin/Debug/DB2Sheet.dll`。
- `PoorMansTSqlFormatterLib35.dll` 已写入 VSTO 清单。
- 对新主程序集执行运行时断言通过：字段列表保持单行，WHERE 条件分行，行尾注释保留，重复格式化稳定。
- Advanced Installer 工程已核对三个新增安装文件的 Component、File 和 MainFeature 记录。
- GitHub CLI 已登录；远程不存在 `v2.0.2`；当前分支为 `main`。
- 测试项目因三个既有编译错误无法生成：`AiTranscriptHtmlTests` 参数类型过期、`AppPresentationTests` 缺少 Ribbon 构造参数、`QueryProfileRepositoryTests` 仍引用已删除的 `TargetSheetName`。本次未顺带修改。
- 尚未在 Excel 或 WPS 宿主中人工查看编辑器格式化效果和模型检测弹窗。

## 阻塞项与下一步

- 无发布配置阻塞。提交当前改动，确保工作区干净后，只运行 `scripts/publish-release.ps1`。
- 发布脚本应完成 Release 构建、Advanced Installer 打包、`v2.0.2` 标签、GitHub Release 和 Gitee 镜像。
- 发布后核验 GitHub、Gitee 的标签、Release 和 `db-to-excel.exe` 附件。