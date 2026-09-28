using System;
using System.Runtime.InteropServices;
using System.Security;
using OfficeCore = Microsoft.Office.Core;

namespace DB2Sheet.UI
{
    /// <summary>向 Excel 提供自定义 Ribbon XML，并把按钮回调转发给加载项入口。</summary>
    /// <remarks>该类型需要对 COM 可见；Ribbon 控件 ID 是 Office 回调标识，不应随展示名称一起修改。</remarks>
    [ComVisible(true)]
    public sealed class DB2SheetRibbon : OfficeCore.IRibbonExtensibility
    {
        private readonly Action _openQueryEditor;
        private readonly Action _openSheetRefresh;
        private readonly Action _openConnections;
        private readonly Action _openSettings;

        /// <summary>创建 Ribbon 并注入各按钮对应的 UI 操作。</summary>
        /// <param name="openQueryEditor">打开 SQL 查询窗体的操作。</param>
        /// <param name="openSheetRefresh">打开批量刷新窗体的操作。</param>
        /// <param name="openConnections">打开连接管理窗体的操作。</param>
        /// <param name="openSettings">打开设置窗体的操作。</param>
        public DB2SheetRibbon(
            Action openQueryEditor,
            Action openSheetRefresh,
            Action openConnections,
            Action openSettings)
        {
            _openQueryEditor = openQueryEditor ?? throw new ArgumentNullException(nameof(openQueryEditor));
            _openSheetRefresh = openSheetRefresh ?? throw new ArgumentNullException(nameof(openSheetRefresh));
            _openConnections = openConnections ?? throw new ArgumentNullException(nameof(openConnections));
            _openSettings = openSettings ?? throw new ArgumentNullException(nameof(openSettings));
        }

        /// <summary>返回 Excel 请求的自定义 Ribbon XML。</summary>
        /// <param name="ribbonId">Office 提供的 Ribbon 标识；当前实现无需区分。</param>
        /// <returns>包含查询、刷新和管理按钮的 XML。</returns>
        public string GetCustomUI(string ribbonId)
        {
            return string.Format(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnLoad"">
  <ribbon>
    <tabs>
      <tab id=""DB2SheetTab"" label=""{0}"">
        <group id=""DB2SheetQueryGroup"" label=""查询"">
          <button id=""DB2SheetQueryButton"" label=""SQL 查询"" size=""large"" imageMso=""DatabaseQueryNew"" onAction=""OpenQueryEditor"" screentip=""打开 SQL 查询与结果预览"" />
          <button id=""DB2SheetRefreshButton"" label=""批量刷新"" size=""large"" imageMso=""RefreshAll"" onAction=""OpenSheetRefresh"" screentip=""从 SQL Sheet 刷新目标工作表"" />
        </group>
        <group id=""DB2SheetManageGroup"" label=""管理"">
          <button id=""DB2SheetConnectionsButton"" label=""连接管理"" imageMso=""Connections"" onAction=""OpenConnections"" />
          <button id=""DB2SheetSettingsButton"" label=""设置"" imageMso=""ControlProperties"" onAction=""OpenSettings"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>", SecurityElement.Escape(AppPresentation.DisplayName));
        }

        /// <summary>接收 Office 创建的 Ribbon UI 对象。</summary>
        /// <param name="ribbonUi">Office Ribbon UI；当前无需缓存。</param>
        public void OnLoad(OfficeCore.IRibbonUI ribbonUi)
        {
        }

        /// <summary>处理“SQL 查询”按钮回调。</summary>
        /// <param name="control">触发回调的 Ribbon 控件。</param>
        public void OpenQueryEditor(OfficeCore.IRibbonControl control)
        {
            _openQueryEditor();
        }

        /// <summary>处理“批量刷新”按钮回调。</summary>
        /// <param name="control">触发回调的 Ribbon 控件。</param>
        public void OpenSheetRefresh(OfficeCore.IRibbonControl control)
        {
            _openSheetRefresh();
        }

        /// <summary>处理“连接管理”按钮回调。</summary>
        /// <param name="control">触发回调的 Ribbon 控件。</param>
        public void OpenConnections(OfficeCore.IRibbonControl control)
        {
            _openConnections();
        }

        /// <summary>处理“设置”按钮回调。</summary>
        /// <param name="control">触发回调的 Ribbon 控件。</param>
        public void OpenSettings(OfficeCore.IRibbonControl control)
        {
            _openSettings();
        }
    }
}
