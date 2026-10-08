using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading;
using DB2Sheet.Models;
using Microsoft.Win32;

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

    /// <summary>解析或搜索本机 java.exe。读取版本会短暂启动 JVM，查找本身不启动。</summary>
    /// <remarks>
    /// 已保存路径、环境变量和常见安装位置属于快速档，供查询启动使用。
    /// 只有「检测 Java」会在快速档落空后扫描 C 盘。32 位进程里的 Program Files 会指向 x86 目录，因此另读 ProgramW6432。
    /// </remarks>
    internal static class JavaRuntimeProbe
    {
        /// <summary>未找到 Java 时引导用户打开 JDBC 环境的说明。</summary>
        public const string MissingJavaMessage =
            "未找到可用的 Java。请打开「JDBC 环境」，点击「检测 Java」，或浏览指定 java.exe。";

        /// <summary>Eclipse Temurin 下载页，供用户自行安装 JDK。</summary>
        public const string DownloadUrl = "https://adoptium.net/temurin/releases/?os=windows&arch=x64&package=jdk";

        /// <summary>快速档和 C 盘都没找到时给用户的说明。</summary>
        private const string NotFoundDetail =
            "本机没有找到 java.exe。请点击「打开 JDK 下载页」安装后再次检测，或浏览指定路径。";

        /// <summary>从 C:\ 向下访问的最深目录层。够到「用户\下载\jdk\bin」这一级。</summary>
        private const int DriveMaxDepth = 5;

        /// <summary>常见安装根向下查找的层数。厂商目录下面通常还有一层 JDK 目录。</summary>
        private const int InstallMaxDepth = 3;

        /// <summary>同一目录文字最少间隔，避免进度日志被每一层刷满。</summary>
        private const int DetailIntervalMilliseconds = 150;

        private static readonly string[] RegistryInstallKeys =
        {
            @"SOFTWARE\JavaSoft\JDK",
            @"SOFTWARE\JavaSoft\Java Development Kit",
            @"SOFTWARE\JavaSoft\JRE",
            @"SOFTWARE\JavaSoft\Java Runtime Environment",
            @"SOFTWARE\Eclipse Adoptium\JDK",
            @"SOFTWARE\AdoptOpenJDK\JDK",
            @"SOFTWARE\Eclipse Foundation\JDK",
            @"SOFTWARE\Amazon Corretto\JDK",
            @"SOFTWARE\Microsoft\JDK",
            @"SOFTWARE\Azul Systems\Zulu",
            @"SOFTWARE\BellSoft\Liberica"
        };

        private static readonly string[] RegistryHomeValueNames = { "JavaHome", "Path", "InstallationPath" };

        private static readonly HashSet<string> SkippedDirectoryNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Windows",
            "$Recycle.Bin",
            "System Volume Information",
            "Recovery",
            "AppData",
            "node_modules",
            ".git",
            "WinSxS",
            "WindowsApps",
            "Temp",
            "Cache"
        };

        /// <summary>按已保存路径解析；为空或无效时走快速档，不扫描 C 盘。</summary>
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

        /// <summary>按快速档搜索本机 java.exe，不扫描整盘。</summary>
        /// <param name="preferredPath">优先尝试的路径，通常是路径框当前值。</param>
        /// <returns>可用结果；都没有时返回失败说明。</returns>
        public static JavaResolution Discover(string preferredPath = null)
        {
            return Search(preferredPath, false, null, CancellationToken.None);
        }

        /// <summary>按阶段搜索 java.exe，并可把当前阶段报告给进度窗。</summary>
        /// <param name="preferredPath">优先尝试的路径。有效时直接采用，不再继续搜索。</param>
        /// <param name="scanDriveC">快速档没有结果时是否扫描 C 盘。查询解析必须传 false。</param>
        /// <param name="progress">进度接收器。为 null 时只搜索不报告。可从后台线程调用。</param>
        /// <param name="cancellationToken">取消标记。取消时抛出 <see cref="OperationCanceledException"/>，已扫到的路径不返回。</param>
        /// <returns>选定的 java.exe；没有可用结果时 Found 为 false。</returns>
        /// <exception cref="OperationCanceledException">调用方取消了搜索。</exception>
        /// <remarks>
        /// 百分比只增不减，且不超过 99。画满并关闭进度窗由操作运行器负责。
        /// 目录拒绝访问时跳过，不把无权限当成检测失败。
        /// </remarks>
        public static JavaResolution Search(
            string preferredPath,
            bool scanDriveC,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken)
        {
            SearchProgress reporter = new SearchProgress(progress);
            List<string> found = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            reporter.Stage(2, "正在检查已填写的路径…");
            cancellationToken.ThrowIfCancellationRequested();
            string preferred = FindInLocation(preferredPath);
            if (preferred != null)
            {
                return JavaResolution.Success(preferred);
            }

            reporter.Stage(8, "正在检查 JAVA_HOME 和 PATH…");
            cancellationToken.ThrowIfCancellationRequested();
            string fromHome = FindInLocation(Environment.GetEnvironmentVariable("JAVA_HOME"));
            if (fromHome != null)
            {
                return JavaResolution.Success(fromHome);
            }

            foreach (string directory in CurrentPathDirectories())
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    string candidate = Path.Combine(directory, "java.exe");
                    if (!File.Exists(candidate)) continue;
                    return JavaResolution.Success(Path.GetFullPath(candidate));
                }
                catch (IOException)
                {
                    // 这一项路径无效时继续看 PATH 里的下一个目录。
                }
                catch (ArgumentException)
                {
                    // 目录含有非法字符时跳过该项。
                }
            }

            reporter.Stage(20, "正在检查注册表中的 Java…");
            CollectFromRegistry(found, seen, cancellationToken);
            string fromRegistry = ChooseBest(found);
            if (fromRegistry != null)
            {
                return JavaResolution.Success(fromRegistry);
            }

            found.Clear();
            seen.Clear();
            reporter.Stage(30, "正在检查常见安装目录…");
            CollectCommonInstalls(found, seen, reporter, cancellationToken);
            string fromCommon = ChooseBest(found);
            if (fromCommon != null)
            {
                return JavaResolution.Success(fromCommon);
            }

            if (!scanDriveC)
            {
                return JavaResolution.Failure(NotFoundDetail);
            }

            found.Clear();
            seen.Clear();
            reporter.Stage(55, "正在扫描 C:\\…");
            CollectDriveC(found, seen, reporter, cancellationToken);
            string fromDrive = ChooseBest(found);
            if (fromDrive != null)
            {
                return JavaResolution.Success(fromDrive);
            }

            return JavaResolution.Failure(NotFoundDetail);
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
        /// <remarks>只向下看一层，返回第一个存在的 java.exe，不比较版本。</remarks>
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
                catch (UnauthorizedAccessException)
                {
                    // 这个根目录拒绝枚举时换下一个，测试目录或安装目录都可能没有权限。
                }
                catch (IOException)
                {
                    // 目录正在变化时跳过该根。
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
                            // 强行结束超时的 java 进程失败时只放弃版本文本，不把清理失败当成检测失败。
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
                // 版本探测是附加信息。java 无法启动时交给调用方显示「无法读取版本」。
                return string.Empty;
            }
        }

        /// <summary>从 32 位和 64 位注册表视图收集安装器写下的 Java 目录。</summary>
        /// <param name="found">收集到的 java.exe。调用方负责清空。</param>
        /// <param name="seen">已收录路径，忽略大小写。</param>
        /// <param name="cancellationToken">取消标记。</param>
        private static void CollectFromRegistry(List<string> found, HashSet<string> seen, CancellationToken cancellationToken)
        {
            RegistryHive[] hives = { RegistryHive.LocalMachine, RegistryHive.CurrentUser };
            RegistryView[] views = { RegistryView.Registry64, RegistryView.Registry32 };
            foreach (RegistryHive hive in hives)
            {
                foreach (RegistryView view in views)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    ReadRegistryView(hive, view, found, seen, cancellationToken);
                }
            }
        }

        /// <summary>打开一个注册表视图并读取已知安装项。打不开时跳过该视图。</summary>
        private static void ReadRegistryView(
            RegistryHive hive,
            RegistryView view,
            List<string> found,
            HashSet<string> seen,
            CancellationToken cancellationToken)
        {
            try
            {
                using (RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, view))
                {
                    foreach (string relative in RegistryInstallKeys)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        using (RegistryKey key = baseKey.OpenSubKey(relative, false))
                        {
                            if (key == null) continue;
                            ReadRegistryTree(key, 0, found, seen, cancellationToken);
                        }
                    }
                }
            }
            catch (IOException)
            {
                // 该视图打不开时还有其它配置单元和位数，不中断检测。
            }
            catch (UnauthorizedAccessException)
            {
                // 没有读取这个视图的权限时跳过。
            }
            catch (SecurityException)
            {
                // 安全策略拒绝打开注册表时跳过该视图。
            }
            catch (ArgumentException)
            {
                // 32 位系统上打开 64 位视图可能被拒绝，换下一个视图。
            }
        }

        /// <summary>读取安装项上的目录值，并有限地向子项查找。安装器常把 JavaHome 写在版本子项里。</summary>
        private static void ReadRegistryTree(
            RegistryKey key,
            int depth,
            List<string> found,
            HashSet<string> seen,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (string valueName in RegistryHomeValueNames)
            {
                AddLocation(ReadRegistryString(key, valueName), found, seen);
            }

            if (depth >= 5) return;
            string[] children = ReadSubKeyNames(key);
            for (int index = 0; index < children.Length; index++)
            {
                using (RegistryKey child = OpenSubKey(key, children[index]))
                {
                    if (child == null) continue;
                    ReadRegistryTree(child, depth + 1, found, seen, cancellationToken);
                }
            }
        }

        /// <summary>读取注册表字符串。值不存在或无权读取时返回 null。</summary>
        private static string ReadRegistryString(RegistryKey key, string name)
        {
            try
            {
                return key.GetValue(name) as string;
            }
            catch (IOException)
            {
                return null;
            }
            catch (UnauthorizedAccessException)
            {
                return null;
            }
            catch (SecurityException)
            {
                return null;
            }
        }

        /// <summary>列出子项名称。列不出时当作没有子项。</summary>
        private static string[] ReadSubKeyNames(RegistryKey key)
        {
            try
            {
                return key.GetSubKeyNames();
            }
            catch (IOException)
            {
                return new string[0];
            }
            catch (UnauthorizedAccessException)
            {
                return new string[0];
            }
            catch (SecurityException)
            {
                return new string[0];
            }
        }

        /// <summary>打开子项。无权打开或项已消失时返回 null。</summary>
        private static RegistryKey OpenSubKey(RegistryKey key, string name)
        {
            try
            {
                return key.OpenSubKey(name, false);
            }
            catch (IOException)
            {
                return null;
            }
            catch (UnauthorizedAccessException)
            {
                return null;
            }
            catch (SecurityException)
            {
                return null;
            }
        }

        /// <summary>搜索 Program Files、用户安装目录，以及名称像 JDK 的顶层文件夹。</summary>
        private static void CollectCommonInstalls(
            List<string> found,
            HashSet<string> seen,
            SearchProgress reporter,
            CancellationToken cancellationToken)
        {
            List<string> roots = ExistingDirectories(CommonInstallRoots());
            for (int index = 0; index < roots.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int percent = SpanPercent(30, 48, index, roots.Count);
                reporter.Stage(percent, "正在检查 " + roots[index] + "…");
                CollectInTree(roots[index], 0, InstallMaxDepth, found, seen, cancellationToken, null);
            }

            reporter.Stage(50, "正在检查名称像 JDK 的目录…");
            CollectNamedInstalls(found, seen, cancellationToken);
        }

        /// <summary>常见安装根。包含 64 位 Program Files，避免 32 位 Excel 只看到 x86 目录。</summary>
        private static IEnumerable<string> CommonInstallRoots()
        {
            yield return Environment.GetEnvironmentVariable("ProgramW6432");
            yield return Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            yield return Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (!string.IsNullOrWhiteSpace(local))
            {
                yield return Path.Combine(local, "Programs");
                yield return Path.Combine(local, "Microsoft", "WinGet", "Packages");
            }

            string profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (!string.IsNullOrWhiteSpace(profile))
            {
                yield return Path.Combine(profile, "scoop", "apps");
            }
        }

        /// <summary>在各固定磁盘和用户目录下，只进入名称像 JDK 的文件夹。</summary>
        private static void CollectNamedInstalls(List<string> found, HashSet<string> seen, CancellationToken cancellationToken)
        {
            foreach (DriveInfo drive in DriveInfo.GetDrives())
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    if (drive.DriveType != DriveType.Fixed || !drive.IsReady) continue;
                    CollectNamedChildren(drive.RootDirectory.FullName, found, seen, cancellationToken);
                }
                catch (IOException)
                {
                    // 某块磁盘暂时不可用时跳过，其它磁盘继续。
                }
                catch (UnauthorizedAccessException)
                {
                    // 磁盘根目录拒绝访问时跳过。
                }
            }

            CollectNamedChildren(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), found, seen, cancellationToken);
            CollectNamedChildren(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), found, seen, cancellationToken);
        }

        /// <summary>进入父目录中名称像 JDK 的直接子文件夹。</summary>
        private static void CollectNamedChildren(
            string parent,
            List<string> found,
            HashSet<string> seen,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(parent)) return;
            ForEachChildDirectory(parent, child =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!NameLooksLikeJdk(Path.GetFileName(child))) return;
                CollectInTree(child, 0, InstallMaxDepth, found, seen, cancellationToken, null);
            });
        }

        /// <summary>扫描 C 盘有限深度，并用顶层目录及其直接子目录推进进度。</summary>
        /// <remarks>跳过 Windows 等系统目录。重解析点只检查其中的 java.exe，不再往下走，避免联接目录循环。</remarks>
        private static void CollectDriveC(
            List<string> found,
            HashSet<string> seen,
            SearchProgress reporter,
            CancellationToken cancellationToken)
        {
            if (!Directory.Exists(@"C:\")) return;
            List<string> tops = new List<string>();
            ForEachChildDirectory(@"C:\", child =>
            {
                if (ShouldSkipDirectory(Path.GetFileName(child))) return;
                tops.Add(child);
            });

            int topTotal = tops.Count;
            for (int index = 0; index < topTotal; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string top = tops[index];
                reporter.Stage(DrivePercent(index, 0, 1, topTotal), "正在扫描 " + top + "…");
                if (IsReparsePoint(top))
                {
                    AddJavaExecutables(top, found, seen);
                    continue;
                }

                AddJavaExecutables(top, found, seen);
                List<string> children = new List<string>();
                ForEachChildDirectory(top, child =>
                {
                    if (ShouldSkipDirectory(Path.GetFileName(child))) return;
                    children.Add(child);
                });

                int childTotal = children.Count;
                if (childTotal == 0)
                {
                    reporter.Detail("正在扫描 " + top + "…");
                    continue;
                }

                for (int childIndex = 0; childIndex < childTotal; childIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    string child = children[childIndex];
                    if (IsReparsePoint(child))
                    {
                        AddJavaExecutables(child, found, seen);
                    }
                    else
                    {
                        // 顶层目录是第 1 层，子目录从第 2 层继续，最深到第 5 层。
                        CollectInTree(child, 2, DriveMaxDepth, found, seen, cancellationToken, directory =>
                            reporter.Detail("正在扫描 " + directory + "…"));
                    }

                    reporter.Report(
                        DrivePercent(index, childIndex + 1, childTotal, topTotal),
                        "正在扫描 " + child + "…",
                        false);
                }
            }
        }

        /// <summary>从目录向下收集 java.exe，直到最大深度。</summary>
        /// <param name="directory">当前目录。</param>
        /// <param name="depth">当前目录相对扫描起点的深度。</param>
        /// <param name="maxDepth">允许访问的最大深度，含当前目录。</param>
        /// <param name="onVisit">每进入一个目录时调用。用于节流后的进度文字，可为 null。</param>
        private static void CollectInTree(
            string directory,
            int depth,
            int maxDepth,
            List<string> found,
            HashSet<string> seen,
            CancellationToken cancellationToken,
            Action<string> onVisit)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (onVisit != null) onVisit(directory);
            AddJavaExecutables(directory, found, seen);
            if (depth >= maxDepth) return;
            ForEachChildDirectory(directory, child =>
            {
                if (ShouldSkipDirectory(Path.GetFileName(child))) return;
                if (IsReparsePoint(child))
                {
                    AddJavaExecutables(child, found, seen);
                    return;
                }

                CollectInTree(child, depth + 1, maxDepth, found, seen, cancellationToken, onVisit);
            });
        }

        /// <summary>检查目录本身和 bin 子目录里的 java.exe。</summary>
        private static void AddJavaExecutables(string directory, List<string> found, HashSet<string> seen)
        {
            AddExecutable(Path.Combine(directory, "java.exe"), found, seen);
            AddExecutable(Path.Combine(directory, "bin", "java.exe"), found, seen);
        }

        /// <summary>把位置解释为 java.exe 或含有 bin\java.exe 的目录后收录。分号分隔的多项会逐项尝试。</summary>
        private static void AddLocation(string location, List<string> found, HashSet<string> seen)
        {
            if (string.IsNullOrWhiteSpace(location)) return;
            string[] parts = location.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int index = 0; index < parts.Length; index++)
            {
                string executable = FindInLocation(parts[index].Trim().Trim('"'));
                if (executable == null) continue;
                if (seen.Add(executable)) found.Add(executable);
            }
        }

        /// <summary>收录一个已存在的 java.exe。路径无法规范化时跳过。</summary>
        private static void AddExecutable(string path, List<string> found, HashSet<string> seen)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
            try
            {
                string full = Path.GetFullPath(path);
                if (seen.Add(full)) found.Add(full);
            }
            catch (IOException)
            {
                // 路径无法规范化时忽略这一条。
            }
            catch (ArgumentException)
            {
                // 含有非法字符时忽略这一条。
            }
        }

        /// <summary>在候选里选择 JDK 8 或更高的最高版本。没有版本文件的排在有版本号的后面。</summary>
        /// <returns>选中的 java.exe；都低于 8 或列表为空时为 null。</returns>
        private static string ChooseBest(List<string> executables)
        {
            JavaCandidate best = null;
            for (int index = 0; index < executables.Count; index++)
            {
                JavaCandidate candidate = Score(executables[index]);
                if (candidate == null) continue;
                if (IsBetter(candidate, best)) best = candidate;
            }

            return best == null ? null : best.Executable;
        }

        /// <summary>根据 release 文件和是否附带 javac 给 java.exe 打分。已知低于 8 的返回 null。</summary>
        private static JavaCandidate Score(string executable)
        {
            if (string.IsNullOrWhiteSpace(executable)) return null;
            int major = ReadMajorVersion(executable);
            if (major > 0 && major < 8) return null;
            string directory = Path.GetDirectoryName(executable);
            bool hasJavac = directory != null && File.Exists(Path.Combine(directory, "javac.exe"));
            return new JavaCandidate(executable, major, hasJavac);
        }

        /// <summary>有版本号的优先于未知版本，版本更高的优先，同版本时优先带 javac 的 JDK。</summary>
        private static bool IsBetter(JavaCandidate candidate, JavaCandidate current)
        {
            if (current == null) return true;
            bool candidateKnown = candidate.Major >= 8;
            bool currentKnown = current.Major >= 8;
            if (candidateKnown != currentKnown) return candidateKnown;
            if (candidate.Major != current.Major) return candidate.Major > current.Major;
            if (candidate.HasJavac != current.HasJavac) return candidate.HasJavac;
            return false;
        }

        /// <summary>从 JDK 根目录的 release 文件解析主版本。1.8 记为 8。读不到时返回 0。</summary>
        private static int ReadMajorVersion(string javaExecutable)
        {
            string release = LocateReleaseFile(javaExecutable);
            if (release == null) return 0;
            try
            {
                foreach (string line in File.ReadLines(release))
                {
                    const string prefix = "JAVA_VERSION=";
                    if (!line.StartsWith(prefix, StringComparison.Ordinal)) continue;
                    string value = line.Substring(prefix.Length).Trim().Trim('"');
                    return ParseMajor(value);
                }
            }
            catch (IOException)
            {
                // 读不到 release 时当作未知版本，仍可入选，但排在有版本号的后面。
            }
            catch (UnauthorizedAccessException)
            {
                // 同上。
            }

            return 0;
        }

        /// <summary>定位 java.exe 旁边或上一级目录中的 release 文件。</summary>
        private static string LocateReleaseFile(string javaExecutable)
        {
            string directory = Path.GetDirectoryName(javaExecutable);
            if (string.IsNullOrEmpty(directory)) return null;
            string beside = Path.Combine(directory, "release");
            if (File.Exists(beside)) return beside;
            string parent = Path.GetDirectoryName(directory);
            if (string.IsNullOrEmpty(parent)) return null;
            string above = Path.Combine(parent, "release");
            return File.Exists(above) ? above : null;
        }

        /// <summary>解析 JAVA_VERSION 的主版本号。</summary>
        private static int ParseMajor(string version)
        {
            if (string.IsNullOrWhiteSpace(version)) return 0;
            string[] parts = version.Split('.');
            if (parts.Length == 0) return 0;
            if (parts[0] == "1" && parts.Length >= 2)
            {
                int legacy;
                return int.TryParse(parts[1], out legacy) ? legacy : 0;
            }

            int modern;
            return int.TryParse(parts[0], out modern) ? modern : 0;
        }

        /// <summary>去掉不存在或重复的目录。</summary>
        private static List<string> ExistingDirectories(IEnumerable<string> paths)
        {
            List<string> result = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in paths)
            {
                if (string.IsNullOrWhiteSpace(path)) continue;
                string full;
                try
                {
                    full = Path.GetFullPath(path);
                }
                catch (ArgumentException)
                {
                    continue;
                }
                catch (IOException)
                {
                    continue;
                }

                if (!seen.Add(full)) continue;
                try
                {
                    if (!Directory.Exists(full)) continue;
                }
                catch (IOException)
                {
                    continue;
                }
                catch (UnauthorizedAccessException)
                {
                    continue;
                }

                result.Add(full);
            }

            return result;
        }

        /// <summary>枚举子目录。这一层拒绝访问或中途失败时停在当前层，不向外抛。</summary>
        private static void ForEachChildDirectory(string directory, Action<string> visit)
        {
            if (visit == null || string.IsNullOrEmpty(directory)) return;
            IEnumerator<string> enumerator = null;
            try
            {
                enumerator = Directory.EnumerateDirectories(directory).GetEnumerator();
                while (true)
                {
                    string child;
                    try
                    {
                        if (!enumerator.MoveNext()) break;
                        child = enumerator.Current;
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // 系统目录经常拒绝枚举。停在这一层，其它分支继续。
                        break;
                    }
                    catch (IOException)
                    {
                        // 目录正在变化或设备不可用时停在这一层。
                        break;
                    }
                    catch (SecurityException)
                    {
                        break;
                    }

                    visit(child);
                }
            }
            catch (UnauthorizedAccessException)
            {
                // 目录本身打不开时跳过。
            }
            catch (IOException)
            {
                // 目录不存在或无法读取时跳过。
            }
            catch (SecurityException)
            {
                // 安全策略拒绝列出目录时跳过。
            }
            finally
            {
                if (enumerator != null) enumerator.Dispose();
            }
        }

        /// <summary>重解析点只用来发现 java.exe，调用方不应再进入。</summary>
        private static bool IsReparsePoint(string path)
        {
            try
            {
                return (File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0;
            }
            catch (IOException)
            {
                // 读不到属性时不要跟着走进去。
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                return true;
            }
        }

        /// <summary>这些目录要么是系统位置，要么会让扫描陷进去，而且不是 JDK 的常规安装点。</summary>
        private static bool ShouldSkipDirectory(string name)
        {
            return string.IsNullOrEmpty(name) || SkippedDirectoryNames.Contains(name);
        }

        /// <summary>文件夹名是否像一套 JDK 或 JRE。</summary>
        private static bool NameLooksLikeJdk(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            return name.IndexOf("jdk", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("jre", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("java", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("temurin", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("corretto", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("zulu", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("liberica", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("graal", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("semeru", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("dragonwell", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("bisheng", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("kona", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>把文件或目录解析成 java.exe 的完整路径。</summary>
        private static string FindInLocation(string location)
        {
            if (string.IsNullOrWhiteSpace(location))
            {
                return null;
            }

            try
            {
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
            catch (IOException)
            {
                // 路径不存在或无法规范化时当作没找到，继续下一处。
                return null;
            }
            catch (ArgumentException)
            {
                // 含有非法字符时当作没找到。
                return null;
            }
            catch (UnauthorizedAccessException)
            {
                // 无权查看该位置时当作没找到。
                return null;
            }
        }

        /// <summary>把序号映射到一段百分比。序号到达总数时落到区间终点。</summary>
        private static int SpanPercent(int start, int end, int index, int count)
        {
            if (count < 1) return start;
            if (index < 0) index = 0;
            if (index > count) index = count;
            int percent = start + (int)((long)(end - start) * index / count);
            if (percent < start) return start;
            if (percent > end) return end;
            return percent;
        }

        /// <summary>C 盘进度落在 55 到 99。分子是已完成的顶层目录加上当前顶层里已完成的直接子目录。</summary>
        private static int DrivePercent(int topIndex, int childDone, int childTotal, int topTotal)
        {
            if (topTotal < 1) return 55;
            if (childTotal < 1) childTotal = 1;
            if (childDone < 0) childDone = 0;
            if (childDone > childTotal) childDone = childTotal;
            if (topIndex < 0) topIndex = 0;
            long numerator = ((long)topIndex * childTotal + childDone) * 44;
            long denominator = (long)topTotal * childTotal;
            int percent = 55 + (int)(numerator / denominator);
            if (percent > 99) return 99;
            if (percent < 55) return 55;
            return percent;
        }

        /// <summary>一次 java.exe 候选及其用于比较的版本信息。</summary>
        private sealed class JavaCandidate
        {
            /// <summary>创建候选。</summary>
            /// <param name="executable">java.exe 完整路径。</param>
            /// <param name="major">主版本。0 表示 release 文件没有给出版本。</param>
            /// <param name="hasJavac">同一目录是否有 javac.exe。</param>
            public JavaCandidate(string executable, int major, bool hasJavac)
            {
                Executable = executable;
                Major = major;
                HasJavac = hasJavac;
            }

            /// <summary>获取 java.exe 路径。</summary>
            public string Executable { get; }

            /// <summary>获取主版本。未知时为 0。</summary>
            public int Major { get; }

            /// <summary>获取是否带有 javac.exe。</summary>
            public bool HasJavac { get; }
        }

        /// <summary>把搜索阶段转成只向前的进度快照。</summary>
        /// <remarks>进度接收器为 null 时所有报告都丢弃，供不需要界面的查询解析使用。</remarks>
        private sealed class SearchProgress
        {
            private readonly IProgress<OperationProgress> _progress;
            private int _shown;
            private int _lastTick;
            private string _lastMessage = string.Empty;

            /// <summary>创建报告器。</summary>
            /// <param name="progress">进度接收器，可为 null。</param>
            public SearchProgress(IProgress<OperationProgress> progress)
            {
                _progress = progress;
            }

            /// <summary>报告一个新阶段。阶段文字立即送出。</summary>
            /// <param name="percent">0 到 99。</param>
            /// <param name="message">显示在进度窗详情和日志中的文字。</param>
            public void Stage(int percent, string message)
            {
                Report(percent, message, true);
            }

            /// <summary>报告当前目录。百分比不变时最多约每 150 毫秒一条。</summary>
            /// <param name="message">当前正在查看的目录。</param>
            public void Detail(string message)
            {
                Report(_shown, message, false);
            }

            /// <summary>送出一条快照。百分比低于已显示值时保持原值。</summary>
            /// <param name="percent">期望百分比，会被限制在已显示值和 99 之间。</param>
            /// <param name="message">详情文字。</param>
            /// <param name="force">为 true 时忽略节流。</param>
            public void Report(int percent, string message, bool force)
            {
                if (_progress == null || string.IsNullOrEmpty(message)) return;
                if (percent < _shown) percent = _shown;
                if (percent > 99) percent = 99;
                int now = Environment.TickCount;
                bool moved = percent > _shown;
                bool same = !moved && string.Equals(message, _lastMessage, StringComparison.Ordinal);
                if (!force && (same || (!moved && unchecked(now - _lastTick) < DetailIntervalMilliseconds))) return;
                _shown = percent;
                _lastMessage = message;
                _lastTick = now;
                _progress.Report(new OperationProgress
                {
                    Stage = OperationStage.Executing,
                    Message = message,
                    Percent = percent,
                    IsIndeterminate = false
                });
            }
        }
    }
}
