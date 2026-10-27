using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Services;

namespace DB2Sheet.UI
{
    /// <summary>记住功能窗体的宽高，以及关闭时是否最大化。</summary>
    /// <remarks>恢复发生在创建句柄之前。句柄已经存在时再切到最大化，会先露出普通尺寸的一帧。</remarks>
    internal static class FormSizeMemory
    {
        /// <summary>立刻恢复尺寸和最大化，并在用户调整或关闭时写回。</summary>
        /// <param name="form">要记住尺寸的窗体。调用时还不能有句柄。</param>
        /// <param name="store">会话设置存储。</param>
        /// <param name="key">该窗体在设置中的稳定名称。</param>
        /// <exception cref="InvalidOperationException">窗体句柄已经创建。</exception>
        public static void Attach(Form form, ISettingsStore store, string key)
        {
            if (form == null) throw new ArgumentNullException(nameof(form));
            if (store == null) throw new ArgumentNullException(nameof(store));
            if (string.IsNullOrWhiteSpace(key)) throw new ArgumentException("窗体尺寸键不能为空。", nameof(key));
            if (form.IsHandleCreated)
                throw new InvalidOperationException("窗体尺寸必须在创建句柄之前恢复，否则最大化会先露出普通尺寸。");

            form.StartPosition = FormStartPosition.Manual;
            Restore(form, store, key);
            form.ResizeEnd += (sender, args) => Save(form, store, key);
            form.FormClosing += (sender, args) => Save(form, store, key);
        }

        /// <summary>返回光标所在屏幕的工作区。不接触窗体，因此不会创建句柄。</summary>
        /// <returns>工作区矩形，已排除任务栏。</returns>
        internal static Rectangle CursorWorkingArea()
        {
            return Screen.FromPoint(Cursor.Position).WorkingArea;
        }

        /// <summary>把窗体宽高收进光标所在屏幕的工作区，使小屏或高缩放下也能看全界面。</summary>
        /// <param name="form">要收敛的窗体。可以还没有句柄。</param>
        /// <param name="margin">四周预留的边距，避免贴着任务栏或屏幕边缘。</param>
        /// <remarks>
        /// 只缩小、不放大：内容本来就放得下时不动。最小化或最大化时跳过，否则会改掉还原尺寸。
        /// 缩小后按「功能窗体相对工作区居中」的约定重新居中，避免窗口缩到左上角。
        /// 低于 <see cref="Form.MinimumSize"/> 的值由 WinForms 自己抬到最小值，这里不再重复判断。
        /// </remarks>
        internal static void ClampToWorkingArea(Form form, int margin = 8)
        {
            if (form == null || form.IsDisposed) return;
            if (form.WindowState != FormWindowState.Normal) return;

            Rectangle area = CursorWorkingArea();
            int width = Math.Min(form.Width, Math.Max(0, area.Width - (margin * 2)));
            int height = Math.Min(form.Height, Math.Max(0, area.Height - (margin * 2)));
            if (width <= 0 || height <= 0) return;
            if (width == form.Width && height == form.Height) return;
            form.Size = new Size(width, height);
            if (form.StartPosition == FormStartPosition.Manual)
            {
                form.Location = new Point(
                    area.Left + Math.Max(0, (area.Width - form.Width) / 2),
                    area.Top + Math.Max(0, (area.Height - form.Height) / 2));
            }
        }

        private static void Restore(Form form, ISettingsStore store, string key)
        {
            if (!TryRead(store, key, out int width, out int height, out bool maximized))
            {
                Center(form);
                return;
            }

            Rectangle area = CursorWorkingArea();
            int minimumWidth = Math.Max(form.MinimumSize.Width, 200);
            int minimumHeight = Math.Max(form.MinimumSize.Height, 150);
            width = Math.Max(minimumWidth, Math.Min(width, area.Width));
            height = Math.Max(minimumHeight, Math.Min(height, area.Height));
            form.Size = new Size(width, height);
            if (maximized)
                form.WindowState = FormWindowState.Maximized;
            else
                Center(form);
        }

        private static void Save(Form form, ISettingsStore store, string key)
        {
            if (form.WindowState == FormWindowState.Minimized) return;
            Rectangle bounds = form.WindowState == FormWindowState.Normal ? form.Bounds : form.RestoreBounds;
            if (bounds.Width <= 0 || bounds.Height <= 0) return;
            bool maximized = form.WindowState == FormWindowState.Maximized;
            string value = bounds.Width.ToString(CultureInfo.InvariantCulture)
                + ","
                + bounds.Height.ToString(CultureInfo.InvariantCulture)
                + (maximized ? ",M" : string.Empty);
            Dictionary<string, string> map = ReadMap(store);
            if (map.TryGetValue(key, out string current) && string.Equals(current, value, StringComparison.Ordinal))
                return;
            map[key] = value;
            store.Set(CoreSettings.FormBounds, WriteMap(map));
        }

        private static void Center(Form form)
        {
            Rectangle area = CursorWorkingArea();
            form.Location = new Point(
                area.Left + Math.Max(0, (area.Width - form.Width) / 2),
                area.Top + Math.Max(0, (area.Height - form.Height) / 2));
        }

        private static bool TryRead(ISettingsStore store, string key, out int width, out int height, out bool maximized)
        {
            width = 0;
            height = 0;
            maximized = false;
            Dictionary<string, string> map = ReadMap(store);
            if (!map.TryGetValue(key, out string value) || string.IsNullOrWhiteSpace(value))
                return false;
            string[] parts = value.Split(',');
            if (parts.Length < 2) return false;
            if (!int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out width)) return false;
            if (!int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out height)) return false;
            maximized = parts.Length > 2 && string.Equals(parts[2], "M", StringComparison.OrdinalIgnoreCase);
            return width > 0 && height > 0;
        }

        private static Dictionary<string, string> ReadMap(ISettingsStore store)
        {
            Dictionary<string, string> map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string raw = store.Get(CoreSettings.FormBounds) ?? string.Empty;
            foreach (string part in raw.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int separator = part.IndexOf('=');
                if (separator <= 0 || separator >= part.Length - 1) continue;
                map[part.Substring(0, separator)] = part.Substring(separator + 1);
            }
            return map;
        }

        private static string WriteMap(Dictionary<string, string> map)
        {
            List<string> parts = new List<string>();
            foreach (KeyValuePair<string, string> pair in map)
                parts.Add(pair.Key + "=" + pair.Value);
            return string.Join(";", parts);
        }
    }
}
