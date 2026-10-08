# DB2Sheet 部署说明

## 部署决策

本项目不使用 VSTO ClickOnce。最终安装包由 Advanced Installer 创建和维护。

发布在仓库根目录运行 `scripts/publish-release.ps1`。脚本不改版本号，不签名，不提交安装工程、证书或构建输出。版本号只改 `Properties/AssemblyInfo.cs`。安装工程和安装包不入库。

脚本按下面的条件执行，任一不满足就停止：

- 工作区没有未提交改动，当前分支是 `main`。
- `AssemblyVersion` 与 `AssemblyFileVersion` 都是三段版本，并且两者相同。
- 本机已有 `installer\db-to-excel.aip`，并且本机安装了该工程所要求版本的 Advanced Installer。
- Release 构建成功。安装工程里指向 `bin\Release` 的文件都存在。
- 安装工程引用 `x64\SQLite.Interop.dll`。Release 已生成该文件但工程未引用时，脚本停止，并且不会改安装工程的文件表。
- 安装工程的 `ProductVersion` 不高于程序集版本。低于程序集版本时，脚本调用 Advanced Installer 对齐版本并生成新的 `ProductCode`。

通过后，脚本用该工程生成 exe，打注解标签 `v` 加三段版本，推送 `main` 和标签，再用 `gh` 创建或更新对应的 GitHub Release 并上传 exe。已有标签必须指向当前提交。

## 目标环境

- Windows 桌面系统
- Microsoft Excel 桌面版
- .NET Framework 4.7.2
- Visual Studio Tools for Office Runtime
- 与目标 Office 位数兼容的数据库本机依赖

## 构建输出

当前 Debug 构建输出：

`bin\Debug\DB2Sheet.dll`

正式安装包应使用 Release 构建，并包含 VSTO 清单、应用程序清单、程序集及运行时依赖。不要只复制 DLL。

## 安装包必须包含的内容

本机安装工程不入库。发布脚本假定 `installer\db-to-excel.aip` 已经存在。安装包需要包含：

- Release 输出、VSTO 清单和应用程序清单。不要只复制 `DB2Sheet.dll`。
- Excel 加载项注册项，加载行为 `LoadBehavior=3`。
- .NET Framework 4.7.2 与 VSTO Runtime 的先决条件检测。
- MySQL、PostgreSQL、SQLite 和 DuckDB 的托管程序集，输出目录中的 `duckdb.dll`，以及 `x64\SQLite.Interop.dll`。
- 不包含 JDK，也不包含 MaxCompute 或其他厂商的 JDBC 驱动。用户不用单独安装 DuckDB。
- 不必单独包含 `db2sheet-jdbc-bridge.jar`。它已嵌入 `DB2Sheet.dll`，首次使用 JDBC 时释放到 `%LocalAppData%\DB2Sheet\db2sheet-jdbc-bridge.jar`。用户通过「JDBC 环境」检测或指定本机 `java.exe`；厂商驱动 jar 在连接方案里选择。

代码签名、Office 位数、安装范围和卸载时是否保留用户数据见「尚待决定」。当前发布脚本不签名。

## 数据和日志

用户配置不放在安装目录，而位于：

`%LocalAppData%\DB2Sheet`

包括：

- `connections.json`：连接方案。密码仍与其他参数一起明文保存。
- `settings.json`
- `jdbc-environment.json`
- `queries\`：每个查询方案一个文件。旧版 `queries.json` 只在首次加载时迁入该目录。
- `Logs\db2sheet-yyyyMMdd.log`

卸载时是否保留这些文件尚未决定，见下文。默认建议保留，避免误删用户配置。

## 安全要求

- 当前发布脚本不签名。生产证书和签名流程见「尚待决定」。临时测试证书不用作生产证书。
- 不在安装包中预置真实数据库凭据。
- 日志目录保持普通用户可写，不要求管理员权限。
- 连接密码仍明文保存在 `connections.json`。Windows DPAPI 或凭据管理器尚未接入。

## Office 位数与依赖

当前发布脚本检查的是 64 位 `x64\SQLite.Interop.dll`。DuckDB 只有 64 位 `duckdb.dll`，32 位 Excel 无法使用该连接。是否同时提供 x86 安装包尚未决定。

`Any CPU` 编译成功并不代表所有本机数据库依赖都能在两种 Office 位数下工作。SQLite 等组件必须在目标机器上实际验证。

## 发布前检查

- [ ] Release 构建成功。
- [ ] 真实 Excel 中 Ribbon 和窗体运行正常。
- [ ] VSTO Runtime 缺失机器上的先决条件安装正常。
- [ ] 普通用户权限可以安装或按设计部署。
- [ ] Excel 信任中心的表现与当时的签名状态一致。当前脚本不签名。
- [ ] SQL Server、MySQL、PostgreSQL、SQLite、DuckDB 依赖加载正常。安装目录中有 `duckdb.dll`、DuckDB 托管程序集，以及 `x64\SQLite.Interop.dll`。
- [ ] 未配置 Java 时，JDBC 连接会提示打开「JDBC 环境」，而不是随安装包附带 JDK 或厂商驱动。
- [ ] 安装后首次启动正常。
- [ ] 升级后，已决定保留的用户配置仍然可用。卸载是否删除 `%LocalAppData%\DB2Sheet` 尚未决定。
- [ ] 卸载后加载项注册项清理。
- [ ] Excel 可正常退出，无残留进程。

## 尚待决定

- 支持的 Office 最低版本。
- x86/x64 支持策略。
- 安装范围：当前用户或所有用户。
- 生产证书和签名流程。
- 用户配置在卸载时的保留策略。
- 自动更新或企业软件分发方式。
