using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace DB2Sheet.Providers
{
    /// <summary>表示一次 Java 可执行文件查找的结果。</summary>
    internal sealed class JavaResolution
    {
        private JavaResolution(string executable, string detail)
        {
            Executable = executable ?? string.Empty;
            Detail = detail ?? string.Empty;
        }

        /// <summary>获取找到的 java.exe 路径；未找到时为空。</summary>
        public string Executable { get; }

        /// <summary>获取未找到时给用户看的说明。</summary>
        public string Detail { get; }

        /// <summary>获取是否已经定位到 java.exe。</summary>
        public bool Found => !string.IsNullOrEmpty(Executable);

        /// <summary>创建成功结果。</summary>
        /// <param name="executable">java.exe 的完整路径。</param>
        /// <returns>成功结果。</returns>
        public static JavaResolution Success(string executable)
        {
            return new JavaResolution(executable, string.Empty);
        }

        /// <summary>创建失败结果。</summary>
        /// <param name="detail">说明文字。</param>
        /// <returns>失败结果。</returns>
        public static JavaResolution Failure(string detail)
        {
            return new JavaResolution(string.Empty, detail);
        }
    }

    /// <summary>解析或搜索本机 java.exe；版本检测会短暂启动 JVM，其它查找不启动。</summary>
    internal static class JavaRuntimeProbe
    {
        /// <summary>未找到 Java 时引导用户打开 JDBC 环境的说明。</summary>
        public const string MissingJavaMessage =
            "未找到可用的 Java。请打开「JDBC 环境」，点击「检测 Java」，或浏览指定 java.exe。";

        /// <summary>Eclipse Temurin 下载页，供用户自行安装 JDK。</summary>
        public const string DownloadUrl = "https://adoptium.net/temurin/releases/?os=windows&arch=x64&package=jdk";

        /// <summary>按已保存路径解析；为空或无效时自动搜索本机。</summary>
        /// <param name="configuredPath">用户保存的 java.exe 或 JDK 目录。</param>
        /// <returns>查找结果。</returns>
        public static JavaResolution Resolve(string configuredPath)
        {
            string configured = (configuredPath ?? string.Empty).Trim();
            if (!string.IsNullOrEmpty(configured))
            {
                string found = FindInLocation(configured);
                if (found != null)
                {
                    return JavaResolution.Success(found);
                }

                // 已保存路径失效时继续搜索，避免用户换机后卡死。
            }

            return Discover(configured);
        }

        /// <summary>按约定顺序搜索本机 java.exe。</summary>
        /// <param name="preferredPath">优先尝试的路径，通常是路径框当前值。</param>
        /// <returns>第一个可用结果；都没有时返回失败说明。</returns>
        public static JavaResolution Discover(string preferredPath = null)
        {
            string preferred = FindInLocation((preferredPath ?? string.Empty).Trim());
            if (preferred != null)
            {
                return JavaResolution.Success(preferred);
            }

            string fromHome = FindInLocation(Environment.GetEnvironmentVariable("JAVA_HOME"));
            if (fromHome != null)
            {
                return JavaResolution.Success(fromHome);
            }

            foreach (string directory in CurrentPathDirectories())
            {
                string candidate = Path.Combine(directory, "java.exe");
                if (File.Exists(candidate))
                {
                    return JavaResolution.Success(Path.GetFullPath(candidate));
                }
            }

            string fromCommon = FindInCommonInstallRoots();
            if (fromCommon != null)
            {
                return JavaResolution.Success(fromCommon);
            }

            return JavaResolution.Failure(
                "本机没有找到 java.exe。请点击「打开 JDK 下载页」安装后再次检测，或浏览指定路径。");
        }

        /// <summary>读取当前进程 PATH 中的目录。</summary>
        /// <returns>目录列表。</returns>
        public static IReadOnlyList<string> CurrentPathDirectories()
        {
            string path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            return path.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim())
                .Where(item => item.Length > 0)
                .ToList()
                .AsReadOnly();
        }

        /// <summary>在给定的常见安装根目录列表中查找 bin\java.exe。供测试注入假目录。</summary>
        /// <param name="roots">安装根目录。</param>
        /// <returns>找到的路径；没有时为 null。</returns>
        public static string FindInInstallRoots(IEnumerable<string> roots)
        {
            if (roots == null) return null;
            foreach (string root in roots)
            {
                if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root)) continue;
                string direct = Path.Combine(root, "bin", "java.exe");
                if (File.Exists(direct))
                {
                    return Path.GetFullPath(direct);
                }

                try
                {
                    foreach (string child in Directory.EnumerateDirectories(root))
                    {
                        string nested = Path.Combine(child, "bin", "java.exe");
                        if (File.Exists(nested))
                        {
                            return Path.GetFullPath(nested);
                        }
                    }
                }
                catch (Exception)
                {
                }
            }

            return null;
        }

        /// <summary>运行 java -version 并返回首行文本。失败时返回空字符串，不抛出。</summary>
        /// <param name="javaExecutable">java.exe 路径。</param>
        /// <returns>版本行；无法启动时为空。</returns>
        public static string TryReadVersion(string javaExecutable)
        {
            if (string.IsNullOrWhiteSpace(javaExecutable) || !File.Exists(javaExecutable))
            {
                return string.Empty;
            }

            try
            {
                ProcessStartInfo start = new ProcessStartInfo
                {
                    FileName = javaExecutable,
                    Arguments = "-version",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using (Process process = Process.Start(start))
                {
                    if (process == null) return string.Empty;
                    string error = process.StandardError.ReadToEnd();
                    string output = process.StandardOutput.ReadToEnd();
                    if (!process.WaitForExit(10000))
                    {
                        try
                        {
                            process.Kill();
                        }
                        catch (Exception)
                        {
                        }

                        return string.Empty;
                    }

                    string text = string.IsNullOrWhiteSpace(error) ? output : error;
                    if (string.IsNullOrWhiteSpace(text)) return string.Empty;
                    string line = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                    return line == null ? string.Empty : line.Trim();
                }
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private static string FindInCommonInstallRoots()
        {
            string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            List<string> roots = new List<string>();
            foreach (string baseDir in new[] { programFiles, programFilesX86 })
            {
                if (string.IsNullOrWhiteSpace(baseDir)) continue;
                roots.Add(Path.Combine(baseDir, "Eclipse Adoptium"));
                roots.Add(Path.Combine(baseDir, "Java"));
                roots.Add(Path.Combine(baseDir, "Microsoft"));
                roots.Add(Path.Combine(baseDir, "Amazon Corretto"));
                roots.Add(Path.Combine(baseDir, "Zulu"));
                roots.Add(Path.Combine(baseDir, "BellSoft"));
            }

            return FindInInstallRoots(roots);
        }

        private static string FindInLocation(string location)
        {
            if (string.IsNullOrWhiteSpace(location))
            {
                return null;
            }

            if (File.Exists(location))
            {
                return Path.GetFullPath(location);
            }

            if (!Directory.Exists(location))
            {
                return null;
            }

            string direct = Path.Combine(location, "java.exe");
            if (File.Exists(direct)) return Path.GetFullPath(direct);
            string nested = Path.Combine(location, "bin", "java.exe");
            if (File.Exists(nested)) return Path.GetFullPath(nested);
            return null;
        }
    }
}
