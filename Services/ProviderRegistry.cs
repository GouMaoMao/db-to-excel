using System;
using System.Collections.Generic;
using System.Linq;
using DB2Sheet.Contracts;

namespace DB2Sheet.Services
{
    /// <summary>以不区分大小写的稳定标识保存数据源提供程序实例。</summary>
    /// <remarks>读取和注册使用同一字典锁，允许跨线程查询；同一标识不允许重复注册。</remarks>
    public sealed class ProviderRegistry : IProviderRegistry
    {
        private readonly Dictionary<string, IDataSourceProvider> _providers =
            new Dictionary<string, IDataSourceProvider>(StringComparer.OrdinalIgnoreCase);

        /// <inheritdoc/>
        public IReadOnlyList<IDataSourceProvider> GetAll()
        {
            lock (_providers)
            {
                return _providers.Values.OrderBy(provider => provider.DisplayName).ToList().AsReadOnly();
            }
        }

        /// <inheritdoc/>
        public IDataSourceProvider GetById(string providerId)
        {
            if (string.IsNullOrWhiteSpace(providerId))
            {
                return null;
            }

            lock (_providers)
            {
                _providers.TryGetValue(providerId, out IDataSourceProvider provider);
                return provider;
            }
        }

        /// <inheritdoc/>
        public void Register(IDataSourceProvider provider)
        {
            if (provider == null)
            {
                throw new ArgumentNullException(nameof(provider));
            }

            lock (_providers)
            {
                if (_providers.ContainsKey(provider.ProviderId))
                {
                    throw new InvalidOperationException("数据源 Provider 已注册：" + provider.ProviderId);
                }

                _providers.Add(provider.ProviderId, provider);
            }
        }
    }
}
