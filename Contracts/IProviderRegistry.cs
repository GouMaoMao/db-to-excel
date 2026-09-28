using System.Collections.Generic;

namespace DB2Sheet.Contracts
{
    /// <summary>保存并按稳定标识查找已启用的数据源提供程序。</summary>
    public interface IProviderRegistry
    {
        /// <summary>获取所有已注册提供程序的只读快照。</summary>
        /// <returns>提供程序列表。</returns>
        IReadOnlyList<IDataSourceProvider> GetAll();
        /// <summary>按内部标识查找提供程序。</summary>
        /// <param name="providerId">提供程序稳定标识。</param>
        /// <returns>匹配的提供程序；不存在时由实现抛出异常。</returns>
        IDataSourceProvider GetById(string providerId);
        /// <summary>注册或替换一个提供程序。</summary>
        /// <param name="provider">要注册的提供程序实例。</param>
        void Register(IDataSourceProvider provider);
    }
}
