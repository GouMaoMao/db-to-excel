using System;
using System.Security.Cryptography;
using System.Text;

namespace DB2Sheet.Services
{
    /// <summary>用当前 Windows 用户的 DPAPI 保护模型密钥。</summary>
    /// <remarks>密文只能由同一用户在同一台电脑解开。不要把明文写入日志或对话文件。</remarks>
    public static class AiSecretProtector
    {
        /// <summary>加密密钥。</summary>
        /// <param name="plain">用户刚输入的密钥。</param>
        /// <returns>可写入 JSON 的 Base64 密文。</returns>
        /// <exception cref="ArgumentException">密钥为空。</exception>
        public static string Protect(string plain)
        {
            if (string.IsNullOrEmpty(plain))
                throw new ArgumentException("密钥不能为空。", nameof(plain));

            byte[] protectedBytes = ProtectedData.Protect(
                Encoding.UTF8.GetBytes(plain),
                null,
                DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(protectedBytes);
        }

        /// <summary>解开先前保存的密钥。</summary>
        /// <param name="protectedText">配置中的 Base64 密文。</param>
        /// <returns>供本次请求使用的明文。调用方不得持久化返回值。</returns>
        /// <exception cref="InvalidOperationException">没有密文，或当前用户解不开。</exception>
        public static string Unprotect(string protectedText)
        {
            if (string.IsNullOrWhiteSpace(protectedText))
                throw new InvalidOperationException("还没有保存密钥。");

            try
            {
                byte[] plain = ProtectedData.Unprotect(
                    Convert.FromBase64String(protectedText),
                    null,
                    DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(plain);
            }
            catch (Exception exception) when (exception is CryptographicException || exception is FormatException)
            {
                throw new InvalidOperationException(
                    "无法读取已保存的密钥。它只能由保存它的 Windows 用户在这台电脑上解开。",
                    exception);
            }
        }
    }
}
