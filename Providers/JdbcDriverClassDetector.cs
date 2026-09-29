using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace DB2Sheet.Providers
{
    /// <summary>从 JDBC 驱动 jar 中读取 META-INF/services/java.sql.Driver 声明的驱动类名。</summary>
    /// <remarks>不启动 JVM。一个 jar 可能声明多个类；调用方决定如何展示给用户。</remarks>
    internal static class JdbcDriverClassDetector
    {
        private const string ServiceEntry = "META-INF/services/java.sql.Driver";

        /// <summary>扫描指定 jar，返回声明的驱动类名（去重、保序）。</summary>
        /// <param name="jarPaths">jar 文件路径。</param>
        /// <returns>类名列表；没有任何声明时为空列表。</returns>
        public static IReadOnlyList<string> Detect(IEnumerable<string> jarPaths)
        {
            List<string> names = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
            if (jarPaths == null) return names.AsReadOnly();

            foreach (string path in jarPaths)
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) continue;
                foreach (string name in ReadServiceNames(path))
                {
                    if (seen.Add(name))
                    {
                        names.Add(name);
                    }
                }
            }

            return names.AsReadOnly();
        }

        private static IEnumerable<string> ReadServiceNames(string jarPath)
        {
            using (FileStream stream = File.OpenRead(jarPath))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                ZipArchiveEntry entry = archive.GetEntry(ServiceEntry)
                    ?? archive.Entries.FirstOrDefault(item =>
                        string.Equals(item.FullName.Replace('\\', '/'), ServiceEntry, StringComparison.OrdinalIgnoreCase));
                if (entry == null)
                {
                    yield break;
                }

                using (Stream content = entry.Open())
                using (StreamReader reader = new StreamReader(content, Encoding.UTF8))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        int comment = line.IndexOf('#');
                        if (comment >= 0) line = line.Substring(0, comment);
                        line = line.Trim();
                        if (line.Length > 0) yield return line;
                    }
                }
            }
        }
    }
}
