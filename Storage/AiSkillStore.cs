using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using DB2Sheet.Infrastructure;
using DB2Sheet.Models;

namespace DB2Sheet.Storage
{
    /// <summary>按一份 Skill 一个文件保存上下文。文件只被读取为文本，不会执行。</summary>
    /// <remarks>单份正文的 UTF-8 字节数不得超过 32KB。</remarks>
    public sealed class AiSkillStore
    {
        /// <summary>正文上限。超过时拒绝保存。</summary>
        public const int MaxBodyBytes = 32 * 1024;

        private readonly object _syncRoot = new object();
        private readonly ApplicationPaths _paths;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        private List<AiSkill> _skills;

        /// <summary>创建存储并读取已有 Skill。</summary>
        /// <param name="paths">本地数据目录。</param>
        public AiSkillStore(ApplicationPaths paths)
        {
            _paths = paths ?? throw new ArgumentNullException(nameof(paths));
            _skills = Load();
        }

        /// <summary>Skill 变化后触发。在锁外调用。</summary>
        public event EventHandler Changed;

        /// <summary>按名称返回 Skill 副本。</summary>
        /// <returns>当前全部 Skill。</returns>
        public IReadOnlyList<AiSkill> GetAll()
        {
            lock (_syncRoot)
            {
                return _skills
                    .OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                    .Select(Copy)
                    .ToList()
                    .AsReadOnly();
            }
        }

        /// <summary>保存 Skill。标识为空时新建。</summary>
        /// <param name="skill">名称和正文都不能为空。</param>
        /// <returns>保存后的副本，含分配好的标识。</returns>
        /// <exception cref="ArgumentException">名称为空、正文为空，或正文超过 32KB。</exception>
        public AiSkill Save(AiSkill skill)
        {
            if (skill == null) throw new ArgumentNullException(nameof(skill));
            if (string.IsNullOrWhiteSpace(skill.Name))
                throw new ArgumentException("Skill 名称不能为空。", nameof(skill));
            if (string.IsNullOrWhiteSpace(skill.Body))
                throw new ArgumentException("Skill 正文不能为空。", nameof(skill));
            if (Encoding.UTF8.GetByteCount(skill.Body) > MaxBodyBytes)
                throw new ArgumentException("Skill 正文超过 32KB。", nameof(skill));

            AiSkill saved;
            lock (_syncRoot)
            {
                saved = Copy(skill);
                if (string.IsNullOrWhiteSpace(saved.Id))
                    saved.Id = Guid.NewGuid().ToString("N");
                saved.Name = saved.Name.Trim();
                saved.UpdatedUtc = DateTime.UtcNow.ToString("o");
                int index = _skills.FindIndex(item =>
                    string.Equals(item.Id, saved.Id, StringComparison.OrdinalIgnoreCase));
                if (index >= 0) _skills[index] = Copy(saved);
                else _skills.Add(Copy(saved));
                AtomicFile.Write(FilePath(saved.Id), _serializer.Serialize(saved));
            }

            Changed?.Invoke(this, EventArgs.Empty);
            return saved;
        }

        /// <summary>删除一份 Skill 文件。</summary>
        /// <param name="id">Skill 标识。</param>
        public void Delete(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return;
            lock (_syncRoot)
            {
                _skills.RemoveAll(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
                string path = FilePath(id);
                if (File.Exists(path))
                    File.Delete(path);
            }

            Changed?.Invoke(this, EventArgs.Empty);
        }

        private List<AiSkill> Load()
        {
            List<AiSkill> skills = new List<AiSkill>();
            if (!Directory.Exists(_paths.AiSkillsDirectory)) return skills;
            foreach (string file in Directory.GetFiles(_paths.AiSkillsDirectory, "*.json"))
            {
                try
                {
                    AiSkill skill = _serializer.Deserialize<AiSkill>(File.ReadAllText(file, Encoding.UTF8));
                    if (skill == null || string.IsNullOrWhiteSpace(skill.Name)) continue;
                    if (string.IsNullOrWhiteSpace(skill.Id))
                        skill.Id = Path.GetFileNameWithoutExtension(file);
                    skills.Add(Copy(skill));
                }
                catch (Exception)
                {
                    AtomicFile.BackupCorrupt(file);
                }
            }

            return skills;
        }

        private string FilePath(string id)
        {
            return Path.Combine(_paths.AiSkillsDirectory, id + ".json");
        }

        private static AiSkill Copy(AiSkill skill)
        {
            if (skill == null) return null;
            return new AiSkill
            {
                Id = skill.Id,
                Name = skill.Name,
                Body = skill.Body,
                UpdatedUtc = skill.UpdatedUtc
            };
        }
    }
}
