using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>加载数据库树彩色图标，并为非活动节点生成同尺寸灰度缓存。</summary>
    /// <remarks>返回的 <see cref="ImageList"/> 及其中图像由调用方随所属控件一起释放。</remarks>
    internal static class DatabaseTreeImageCatalog
    {
        private const string DatabaseKey = "database";
        private const string TableKey = "table";
        private const string ViewKey = "view";

        private static readonly IReadOnlyDictionary<string, string> ProviderResources =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["sqlserver"] = "DB_sqlserver.png",
                ["mysql"] = "DB_Mysql.png",
                ["postgresql"] = "DB_PostgreSQL.png",
                ["sqlite"] = "DB_sqlite.png"
            };

        /// <summary>创建包含全部彩色和灰度图标的 16 像素图像列表。</summary>
        /// <returns>由调用方负责释放的图像列表。</returns>
        /// <exception cref="InvalidOperationException">约定的嵌入图片缺失或无法读取。</exception>
        public static ImageList CreateImageList()
        {
            ImageList images = new ImageList
            {
                ColorDepth = ColorDepth.Depth32Bit,
                ImageSize = new Size(16, 16),
                TransparentColor = Color.Transparent
            };
            try
            {
                foreach (KeyValuePair<string, string> provider in ProviderResources)
                    AddPair(images, "provider-" + provider.Key, provider.Value);
                AddPair(images, DatabaseKey, "icon_db.png");
                AddPair(images, TableKey, "icon_table.png");
                AddPair(images, ViewKey, "icon_view.png");
                return images;
            }
            catch
            {
                images.Dispose();
                throw;
            }
        }

        /// <summary>获取数据库类型连接节点的图像键。没有专用图标的类型使用数据库图标。</summary>
        public static string Provider(string providerId, bool active)
        {
            string id = (providerId ?? string.Empty).ToLowerInvariant();
            if (!ProviderResources.ContainsKey(id)) return Database(active);
            return Key("provider-" + id, active);
        }

        /// <summary>获取数据库节点的图像键。</summary>
        public static string Database(bool active) => Key(DatabaseKey, active);

        /// <summary>获取表节点的图像键。</summary>
        public static string Table(bool active) => Key(TableKey, active);

        /// <summary>获取视图节点的图像键。</summary>
        public static string View(bool active) => Key(ViewKey, active);

        private static string Key(string baseKey, bool active) => baseKey + (active ? "-color" : "-gray");

        private static void AddPair(ImageList images, string key, string fileName)
        {
            using (Image source = Load(fileName))
            using (Bitmap color = Resize(source))
            using (Bitmap gray = ToGray(color))
            {
                images.Images.Add(Key(key, true), (Image)color.Clone());
                images.Images.Add(Key(key, false), (Image)gray.Clone());
            }
        }

        private static Image Load(string fileName)
        {
            string resourceName = "DB2Sheet.Resources." + fileName;
            Assembly assembly = typeof(DatabaseTreeImageCatalog).Assembly;
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    throw new InvalidOperationException("找不到数据库树嵌入图标资源：Resources/" + fileName);
                using (Image image = Image.FromStream(stream))
                    return (Image)image.Clone();
            }
        }

        private static Bitmap Resize(Image source)
        {
            Bitmap result = new Bitmap(16, 16, PixelFormat.Format32bppArgb);
            using (Graphics graphics = Graphics.FromImage(result))
            {
                graphics.Clear(Color.Transparent);
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphics.DrawImage(source, new Rectangle(0, 0, 16, 16));
            }
            return result;
        }

        private static Bitmap ToGray(Image source)
        {
            Bitmap result = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
            ColorMatrix matrix = new ColorMatrix(new[]
            {
                new[] { 0.30F, 0.30F, 0.30F, 0F, 0F },
                new[] { 0.59F, 0.59F, 0.59F, 0F, 0F },
                new[] { 0.11F, 0.11F, 0.11F, 0F, 0F },
                new[] { 0F, 0F, 0F, 1F, 0F },
                new[] { 0F, 0F, 0F, 0F, 1F }
            });
            using (Graphics graphics = Graphics.FromImage(result))
            using (ImageAttributes attributes = new ImageAttributes())
            {
                attributes.SetColorMatrix(matrix);
                graphics.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height), 0, 0, source.Width, source.Height, GraphicsUnit.Pixel, attributes);
            }
            return result;
        }
    }
}
