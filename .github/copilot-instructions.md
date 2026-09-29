# Copilot Instructions

## 项目指南
- 该项目不使用 VSTO ClickOnce 部署，后续由 Advanced Installer 封装和部署。

## 代码文档规范
- 后续新增或修改的非生成代码必须同步编写中文 XML 文档注释；自动生成的 `*.Designer.cs` 等文件不手动修改。
- 每个类、接口、枚举和结构都应使用 `<summary>` 说明职责和设计目的。
- 关键公开成员应根据需要使用 `<summary>`、`<param>`、`<returns>`、`<remarks>` 和 `<exception>`，说明主要输入、输出及异常条件。
- 涉及文件、数据库、Excel COM、异步任务或 UI 的代码，应明确说明副作用、线程要求、取消行为、资源所有权和释放责任。
- 注释应解释设计意图和特殊注意事项，避免只复述代码语句。

## DB2Sheet 进度窗口
- DB2Sheet 的进度窗口在任务成功完成后应自动关闭；仅在发生错误时保留窗口供用户查看，不要求用户手动点击关闭。
- DB2Sheet 的所有应用窗体应显示项目自定义 rocket.ico 图标，而不是 WinForms 默认图标。
- DB2Sheet 的所有 WinForms 窗体应使用比宋体更美观的统一中文界面字体，优先使用“Microsoft YaHei UI（微软雅黑 UI）”。
- 查询方案信息必须显示在左下角的查询方案区域，不要放在底部状态栏。