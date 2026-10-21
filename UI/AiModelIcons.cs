using System;
using System.Collections.Generic;
using System.Drawing;

namespace DB2Sheet.UI
{
    /// <summary>按显示名和型号选出厂商图标。对不上时用机器人。</summary>
    /// <remarks>图像按进程缓存，按钮和菜单共用，不在窗体关闭时释放。</remarks>
    internal static class AiModelIcons
    {
        private static readonly object Gate = new object();
        private static readonly Dictionary<string, Image> Cache =
            new Dictionary<string, Image>(StringComparer.OrdinalIgnoreCase);

        private static readonly string[][] Vendors =
        {
            new[] { "deepseek", "llm_deepseek.png" },
            new[] { "claude", "llm_claude.png" },
            new[] { "anthropic", "llm_claude.png" },
            new[] { "qwen", "llm_qwen.png" },
            new[] { "通义", "llm_qwen.png" },
            new[] { "千问", "llm_qwen.png" },
            new[] { "minimax", "llm_minimax.png" },
            new[] { "kimi", "llm_kimi.png" },
            new[] { "moonshot", "llm_kimi.png" },
            new[] { "gemini", "llm_gemini.png" },
            new[] { "zhipu", "llm_zhipu.png" },
            new[] { "智谱", "llm_zhipu.png" },
            new[] { "glm", "llm_zhipu.png" },
            new[] { "doubao", "llm_doubao.png" },
            new[] { "豆包", "llm_doubao.png" },
            new[] { "openai", "llm_openai.png" },
            new[] { "gpt", "llm_openai.png" }
        };

        /// <summary>取 16 像素厂商图标。</summary>
        /// <param name="displayName">配置显示名。</param>
        /// <param name="modelIds">该配置下的型号。用来在显示名对不上时再认一次。</param>
        /// <returns>缓存中的位图。</returns>
        public static Image For(string displayName, IEnumerable<string> modelIds)
        {
            if (modelIds != null)
            {
                foreach (string id in modelIds)
                {
                    string file = Match(id);
                    if (file != null)
                        return Load(file, false);
                }
            }

            string named = Match(displayName);
            return Load(named ?? "icon_robot.png", false);
        }

        /// <summary>Skill 图标。未选用时为灰色。</summary>
        /// <param name="active">正在使用某个 Skill 时为 true。</param>
        /// <returns>缓存中的 16 像素位图。</returns>
        public static Image Skill(bool active)
        {
            return Load("icon_skills.png", !active);
        }

        private static string Match(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            foreach (string[] vendor in Vendors)
            {
                if (text.IndexOf(vendor[0], StringComparison.OrdinalIgnoreCase) >= 0)
                    return vendor[1];
            }

            return null;
        }

        private static Image Load(string fileName, bool grayscale)
        {
            string key = fileName + (grayscale ? "|gray" : "|color");
            lock (Gate)
            {
                Image cached;
                if (Cache.TryGetValue(key, out cached))
                    return cached;
                Image created = DatabaseTreeImageCatalog.LoadActionIcon(fileName, 16, grayscale);
                Cache[key] = created;
                return created;
            }
        }
    }
}
