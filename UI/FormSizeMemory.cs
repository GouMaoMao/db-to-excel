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
    internal static class FormSizeMemory
    {
        /// <summary>在窗体加载时恢复尺寸，并在用户调整或关闭时写回。</summary>
        /// <param name="form">要记住尺寸的窗体。</param>
        /// <param name="store">会话设置存储。</param>
        /// <param name="key">该窗体在设置中的稳定名称。</param>
        public static void Attach(Form form, ISettingsStore store, string key)
        {
            if (form == null) throw new ArgumentNullException(nameof(form));
            if (store == null) throw new ArgumentNullException(nameof(store));
            if (string.IsNullOrWhiteSpace(key)) throw new ArgumentException("窗体尺寸键不能为空。", nameof(key));

            form.StartPosition = FormStartPosition.Manual;
            form.Load += (sender, args) => Restore(form, store, key);
            form.ResizeEnd += (sender, args) => Save(form, store, key);
            form.FormClosing += (sender, args) => Save(form, store, key);
        }

        private static void Restore(Form form, ISettingsStore store, string key)
        {
            if (!TryRead(store, key, out int width, out int height, out bool maximized))
            {
                Center(form);
                return;
            }

            Rectangle area = Screen.FromControl(form).WorkingArea;
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
            Rectangle area = Screen.FromControl(form).WorkingArea;
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
