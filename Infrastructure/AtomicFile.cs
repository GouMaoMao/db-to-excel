using System;
using System.IO;
using System.Text;

namespace DB2Sheet.Infrastructure
{
    /// <summary>用临时文件替换的方式写入文本，避免写到一半留下损坏文件。</summary>
    internal static class AtomicFile
    {
        /// <summary>把文本写入目标路径。目标已存在时先写到临时文件再替换。</summary>
        /// <param name="path">目标文件路径。所在目录不存在时会创建。</param>
        /// <param name="content">完整文件内容。</param>
        /// <remarks>替换失败时临时文件可能留下。调用方负责让用户看到异常，不要吞掉。</remarks>
        public static void Write(string path, string content)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("文件路径不能为空。", nameof(path));
            string directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            string temporary = path + ".tmp";
            File.WriteAllText(temporary, content ?? string.Empty, Encoding.UTF8);
            if (!File.Exists(path))
            {
                File.Move(temporary, path);
                return;
            }

            string backup = path + ".bak";
            if (File.Exists(backup))
                File.Delete(backup);
            File.Replace(temporary, path, backup, true);
        }

        /// <summary>把读失败的文件改名为带时间的备份，避免下次启动再次读到同一份损坏内容。</summary>
        /// <param name="path">损坏文件路径。不存在时不做任何事。</param>
        public static void BackupCorrupt(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
            string backup = path + ".corrupt-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss");
            File.Move(path, backup);
        }
    }
}
