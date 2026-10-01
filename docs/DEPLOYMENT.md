# DB2Sheet 部署说明

## 部署决策

本项目不使用 VSTO ClickOnce。最终安装包由 Advanced Installer 创建和维护。

当前文档记录部署边界和待办事项；在真实 Excel 集成测试完成前，不应视为最终发布流程。

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

## Advanced Installer 待配置项

1. 创建与 Office/VSTO 加载项兼容的安装项目。
2. 选择与目标 Office 位数一致的安装包策略。
3. 添加 Release 输出和 VSTO 清单。
4. 注册 Excel 加载项所需的 HKCU 或 HKLM 注册表项。
5. 设置加载行为 `LoadBehavior=3`。
6. 配置受信任发布者和代码签名证书。
7. 检测并按需安装 .NET Framework 4.7.2 与 VSTO Runtime。
8. 包含 MySQL、PostgreSQL、SQLite 和 DuckDB 的托管程序集，以及输出目录中的 `duckdb.dll`。不要包含 JDK，也不要包含 MaxCompute 或其他厂商的 JDBC 驱动。用户不用单独安装 DuckDB。
9. 不必单独包含 `db2sheet-jdbc-bridge.jar`。它已嵌入 `DB2Sheet.dll`，运行时释放到 `%LocalAppData%\DB2Sheet`。用户可通过「JDBC 环境」检测或指定本机 JDK；厂商 JDBC 驱动 jar 在新建连接时选择，不要打进安装包。
10. 配置升级代码、产品版本和卸载行为。
11. 验证安装、修复、升级和卸载后 Excel 状态。

## 数据和日志

用户配置不放在安装目录，而位于：

`%LocalAppData%\DB2Sheet`

包括连接、查询、设置和日志。卸载时是否保留用户数据应作为安装选项明确说明，默认建议保留，避免误删用户配置。

## 安全要求

- Release 安装包和 VSTO 清单需要使用有效代码签名证书。
- 不使用临时测试证书作为生产发布证书。
- 不在安装包中预置真实数据库凭据。
- 日志目录应保持普通用户可写，不要求管理员权限。
- 生产部署前应评估连接密码使用 Windows DPAPI 或凭据管理器加密。

## Office 位数与依赖

需要明确支持范围：

- 仅 x64 Office；或
- 同时提供 x86 和 x64 安装包。

SQLite 等组件可能包含位数相关本机库，必须在目标机器上实际验证。`Any CPU` 编译成功并不代表所有本机数据库依赖都能在两种 Office 位数下工作。DuckDB 只有 64 位 `duckdb.dll`，32 位 Excel 无法使用该连接。

## 发布前检查

- [ ] Release 构建成功。
- [ ] 真实 Excel 中 Ribbon 和窗体运行正常。
- [ ] VSTO Runtime 缺失机器上的先决条件安装正常。
- [ ] 普通用户权限可以安装或按设计部署。
- [ ] Excel 信任中心不会阻止已签名加载项。
- [ ] SQL Server、MySQL、PostgreSQL、SQLite、DuckDB 依赖加载正常。安装目录中有 `duckdb.dll` 以及 DuckDB 托管程序集。
- [ ] 未配置 Java 时，JDBC 连接会提示打开「JDBC 环境」，而不是随安装包附带 JDK 或厂商驱动。
- [ ] 安装后首次启动正常。
- [ ] 升级后用户配置保留。
- [ ] 卸载后加载项注册项清理。
- [ ] Excel 可正常退出，无残留进程。

## 尚待决定

- 支持的 Office 最低版本。
- x86/x64 支持策略。
- 安装范围：当前用户或所有用户。
- 生产证书和签名流程。
- 用户配置在卸载时的保留策略。
- 自动更新或企业软件分发方式。
