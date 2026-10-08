using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security;
using OfficeCore = Microsoft.Office.Core;

namespace DB2Sheet.UI
{
    /// <summary>向 Excel 提供自定义 Ribbon XML，并把按钮回调转发给加载项入口。</summary>
    /// <remarks>
    /// 该类型需要对 COM 可见；Ribbon 控件 ID 是 Office 回调标识，不应随展示名称一起修改。
    /// 按钮图标在首次绘制时从嵌入资源生成，并缓存为带透明通道的 OLE 图片直到加载项卸载。
    /// </remarks>
    [ComVisible(true)]
    public sealed class DB2SheetRibbon : OfficeCore.IRibbonExtensibility
    {
        private readonly Dictionary<string, stdole.IPictureDisp> _pictures =
            new Dictionary<string, stdole.IPictureDisp>(StringComparer.Ordinal);

        private readonly Action _openQueryEditor;
        private readonly Action _openSheetRefresh;
        private readonly Action _openConnections;
        private readonly Action _openSettings;
        private readonly Action _openJdbcEnvironment;

        /// <summary>创建 Ribbon 并注入各按钮对应的 UI 操作。</summary>
        /// <param name="openQueryEditor">打开 SQL 查询窗体的操作。</param>
        /// <param name="openSheetRefresh">打开批量刷新窗体的操作。</param>
        /// <param name="openConnections">打开连接管理窗体的操作。</param>
        /// <param name="openSettings">打开设置窗体的操作。</param>
        /// <param name="openJdbcEnvironment">打开 JDBC 环境窗体的操作。</param>
        public DB2SheetRibbon(
            Action openQueryEditor,
            Action openSheetRefresh,
            Action openConnections,
            Action openSettings,
            Action openJdbcEnvironment)
        {
            _openQueryEditor = openQueryEditor ?? throw new ArgumentNullException(nameof(openQueryEditor));
            _openSheetRefresh = openSheetRefresh ?? throw new ArgumentNullException(nameof(openSheetRefresh));
            _openConnections = openConnections ?? throw new ArgumentNullException(nameof(openConnections));
            _openSettings = openSettings ?? throw new ArgumentNullException(nameof(openSettings));
            _openJdbcEnvironment = openJdbcEnvironment ?? throw new ArgumentNullException(nameof(openJdbcEnvironment));
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
          <button id=""DB2SheetQueryButton"" label=""SQL 查询"" size=""large"" getImage=""GetImage"" onAction=""OpenQueryEditor"" screentip=""打开 SQL 查询与结果预览"" />
          <button id=""DB2SheetRefreshButton"" label=""批量刷新"" size=""large"" getImage=""GetImage"" onAction=""OpenSheetRefresh"" screentip=""从 SQL Sheet 刷新目标工作表"" />
        </group>
        <group id=""DB2SheetManageGroup"" label=""管理"">
          <button id=""DB2SheetConnectionsButton"" label=""连接管理"" getImage=""GetImage"" onAction=""OpenConnections"" />
          <menu id=""DB2SheetSettingsMenu"" label=""设置"" getImage=""GetImage"">
            <button id=""DB2SheetSettingsButton"" label=""参数设置"" getImage=""GetImage"" onAction=""OpenSettings"" />
            <button id=""DB2SheetJdbcEnvironmentButton"" label=""JDBC 环境"" getImage=""GetImage"" onAction=""OpenJdbcEnvironment"" />
          </menu>
        </group>
        <group id=""DB2SheetAboutGroup"" label=""关于"">
          <button id=""DB2SheetVersionLabel"" getLabel=""GetAboutLabel"" getImage=""GetImage"" onAction=""ShowVersion"" screentip=""当前插件版本"" />
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

        /// <summary>处理“参数设置”按钮回调。</summary>
        /// <param name="control">触发回调的 Ribbon 控件。</param>
        public void OpenSettings(OfficeCore.IRibbonControl control)
        {
            _openSettings();
        }

        /// <summary>处理“JDBC 环境”按钮回调。</summary>
        /// <param name="control">触发回调的 Ribbon 控件。</param>
        public void OpenJdbcEnvironment(OfficeCore.IRibbonControl control)
        {
            _openJdbcEnvironment();
        }

        /// <summary>按控件 ID 返回嵌入图标。由 Office 在绘制 Ribbon 时于 UI 线程调用。</summary>
        /// <param name="control">正在绘制的按钮或菜单。未知 ID 不显示图标。</param>
        /// <returns>缓存的 OLE 图片。同一文件和尺寸只解码一次，由本对象持有到加载项卸载。</returns>
        /// <exception cref="InvalidOperationException">约定的嵌入图片缺失、无法读取，或无法交给 Office。</exception>
        public stdole.IPictureDisp GetImage(OfficeCore.IRibbonControl control)
        {
            string fileName;
            int size;
            if (!TryResolveIcon(control == null ? null : control.Id, out fileName, out size))
                return null;

            string cacheKey = fileName + "|" + size.ToString(CultureInfo.InvariantCulture);
            stdole.IPictureDisp picture;
            if (_pictures.TryGetValue(cacheKey, out picture))
                return picture;

            using (Bitmap image = (Bitmap)DatabaseTreeImageCatalog.LoadActionIcon(fileName, size))
                picture = RibbonPicture.FromBitmap(image);

            _pictures.Add(cacheKey, picture);
            return picture;
        }

        /// <summary>把功能区控件 ID 映射到嵌入图标和 Office 建议尺寸。</summary>
        /// <param name="controlId">Ribbon 控件 ID。</param>
        /// <param name="fileName">Resources 目录下的文件名。</param>
        /// <param name="size">大按钮为 32，小按钮和菜单项为 16。</param>
        /// <returns>该控件需要自定义图标时为 true。</returns>
        private static bool TryResolveIcon(string controlId, out string fileName, out int size)
        {
            fileName = null;
            size = 16;
            switch (controlId)
            {
                case "DB2SheetQueryButton":
                    fileName = "icon_code.png";
                    size = 32;
                    return true;
                case "DB2SheetRefreshButton":
                    fileName = "icon_piliangzhixing.png";
                    size = 32;
                    return true;
                case "DB2SheetConnectionsButton":
                    fileName = "icon_connect.png";
                    return true;
                case "DB2SheetSettingsMenu":
                case "DB2SheetSettingsButton":
                    fileName = "icon_setting.png";
                    return true;
                case "DB2SheetJdbcEnvironmentButton":
                    fileName = "icon_java.png";
                    return true;
                case "DB2SheetVersionLabel":
                    fileName = "icon_info.png";
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>版本按钮只展示图标和版本文案，点击不打开窗体。</summary>
        /// <param name="control">版本按钮。回调存在是为了避免点击时 Office 提示缺少处理程序。</param>
        public void ShowVersion(OfficeCore.IRibbonControl control)
        {
        }

        /// <summary>返回「关于」组的版本标签。由 Office 在绘制 Ribbon 时调用。</summary>
        /// <param name="control">版本标签控件；Id 不是版本标签时返回空字符串。</param>
        /// <returns>形如“版本 1.0.1”的文本；未知控件返回空字符串。</returns>
        public string GetAboutLabel(OfficeCore.IRibbonControl control)
        {
            return FormatAboutLabel(control == null ? null : control.Id);
        }

        /// <summary>把关于组控件 ID 映射为版本文案，供 Ribbon 回调和测试共用。</summary>
        /// <param name="controlId">Ribbon 控件 ID。</param>
        /// <returns>版本标签文本；未知 ID 返回空字符串。</returns>
        internal static string FormatAboutLabel(string controlId)
        {
            if (controlId == "DB2SheetVersionLabel")
                return "版本 " + AppPresentation.FileVersion;

            return string.Empty;
        }

        /// <summary>把带透明通道的位图交给 Office，避免 GDI+ 的 GetHbitmap 把透明像素涂成底色。</summary>
        /// <remarks>生成自下而上的 32 位 DIB，像素保持 PNG 的直通 Alpha。OLE 图片拥有该位图并负责释放。</remarks>
        private static class RibbonPicture
        {
            private const int BitmapType = 1;
            private const int RgbColors = 0;

            /// <summary>复制位图。调用方可随即释放源图。</summary>
            /// <param name="bitmap">32 位或可转换为 32 位的源图。</param>
            /// <returns>功能区可直接绘制的图片。</returns>
            /// <exception cref="InvalidOperationException">无法创建带透明通道的位图。</exception>
            public static stdole.IPictureDisp FromBitmap(Bitmap bitmap)
            {
                IntPtr dib = CreateAlphaBitmap(bitmap);
                try
                {
                    return CreatePicture(dib);
                }
                catch
                {
                    DeleteObject(dib);
                    throw;
                }
            }

            private static stdole.IPictureDisp CreatePicture(IntPtr dib)
            {
                PictureDescription description = new PictureDescription
                {
                    Size = Marshal.SizeOf(typeof(PictureDescription)),
                    Type = BitmapType,
                    BitmapHandle = dib
                };
                Guid pictureId = typeof(stdole.IPictureDisp).GUID;
                return OleCreatePictureIndirect(ref description, ref pictureId, true);
            }

            private static IntPtr CreateAlphaBitmap(Bitmap bitmap)
            {
                int width = bitmap.Width;
                int height = bitmap.Height;
                Rectangle bounds = new Rectangle(0, 0, width, height);
                BitmapData data = bitmap.LockBits(bounds, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                try
                {
                    BitmapInfoHeader header = new BitmapInfoHeader
                    {
                        Size = Marshal.SizeOf(typeof(BitmapInfoHeader)),
                        Width = width,
                        Height = height,
                        Planes = 1,
                        BitCount = 32
                    };
                    IntPtr bits;
                    IntPtr dib = CreateDIBSection(IntPtr.Zero, ref header, RgbColors, out bits, IntPtr.Zero, 0);
                    if (dib == IntPtr.Zero)
                        throw new InvalidOperationException("无法创建功能区图标。");

                    try
                    {
                        int rowBytes = width * 4;
                        byte[] row = new byte[rowBytes];
                        for (int y = 0; y < height; y++)
                        {
                            Marshal.Copy(IntPtr.Add(data.Scan0, y * data.Stride), row, 0, rowBytes);
                            Marshal.Copy(row, 0, IntPtr.Add(bits, (height - 1 - y) * rowBytes), rowBytes);
                        }

                        return dib;
                    }
                    catch
                    {
                        DeleteObject(dib);
                        throw;
                    }
                }
                finally
                {
                    bitmap.UnlockBits(data);
                }
            }

            [DllImport("gdi32.dll", SetLastError = true)]
            private static extern IntPtr CreateDIBSection(
                IntPtr hdc,
                ref BitmapInfoHeader header,
                uint usage,
                out IntPtr bits,
                IntPtr section,
                uint offset);

            [DllImport("gdi32.dll")]
            private static extern bool DeleteObject(IntPtr handle);

            [DllImport("oleaut32.dll", PreserveSig = false)]
            private static extern stdole.IPictureDisp OleCreatePictureIndirect(
                ref PictureDescription description,
                ref Guid pictureId,
                bool ownBitmap);

            [StructLayout(LayoutKind.Sequential)]
            private struct BitmapInfoHeader
            {
                public int Size;
                public int Width;
                public int Height;
                public short Planes;
                public short BitCount;
                public int Compression;
                public int ImageSize;
                public int XPixelsPerMeter;
                public int YPixelsPerMeter;
                public int ColorsUsed;
                public int ColorsImportant;
            }

            [StructLayout(LayoutKind.Sequential)]
            private struct PictureDescription
            {
                public int Size;
                public int Type;
                public IntPtr BitmapHandle;
                public IntPtr Palette;
            }
        }
    }
}
