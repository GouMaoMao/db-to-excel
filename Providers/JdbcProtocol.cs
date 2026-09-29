using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DB2Sheet.Providers
{
    /// <summary>描述发往 JDBC 转接进程的一条请求。</summary>
    /// <remarks>密码只进入标准输入中的 JSON，调用方不得把序列化结果写入日志。</remarks>
    internal sealed class JdbcBridgeRequest
    {
        /// <summary>获取或设置请求标识。取消操作使用独立标识，避免打断查询结果流。</summary>
        public string Id { get; set; }

        /// <summary>获取或设置操作名，例如 query、test、cancel。</summary>
        public string Op { get; set; }

        /// <summary>获取或设置要取消的查询标识。</summary>
        public string QueryId { get; set; }

        /// <summary>获取或设置 JDBC URL。</summary>
        public string Url { get; set; }

        /// <summary>获取或设置用户名或 AccessKey ID。</summary>
        public string User { get; set; }

        /// <summary>获取或设置密码或 AccessKey Secret。</summary>
        public string Password { get; set; }

        /// <summary>获取或设置 JDBC 驱动类名。</summary>
        public string Driver { get; set; }

        /// <summary>获取或设置本连接使用的驱动 jar 绝对路径。</summary>
        public List<string> Jars { get; set; }

        /// <summary>获取或设置 SQL 文本。</summary>
        public string Sql { get; set; }

        /// <summary>获取或设置目录或架构名称。</summary>
        public string Catalog { get; set; }

        /// <summary>获取或设置命令超时秒数。</summary>
        public int TimeoutSeconds { get; set; }
    }

    /// <summary>描述转接进程返回的一帧消息。</summary>
    internal sealed class JdbcBridgeReply
    {
        /// <summary>获取或设置对应的请求标识。</summary>
        public string Id { get; set; }

        /// <summary>获取或设置消息类型，例如 columns、rows、end、error。</summary>
        public string Type { get; set; }

        /// <summary>获取或设置错误或状态文本。</summary>
        public string Message { get; set; }

        /// <summary>获取或设置结果列。</summary>
        public List<JdbcColumnInfo> Columns { get; set; }

        /// <summary>获取或设置一批结果行。单元格在按列类型转换前保留 JSON 原始值。</summary>
        public List<object[]> Rows { get; set; }

        /// <summary>获取或设置目录名称。</summary>
        public List<string> Names { get; set; }

        /// <summary>获取或设置表和视图。</summary>
        public List<JdbcObjectInfo> Objects { get; set; }
    }

    /// <summary>描述 JDBC 结果中的一列及其约定的 CLR 类型短名。</summary>
    internal sealed class JdbcColumnInfo
    {
        /// <summary>获取或设置列名。</summary>
        public string Name { get; set; }

        /// <summary>获取或设置类型短名，例如 string、long、decimal。</summary>
        public string Type { get; set; }
    }

    /// <summary>描述 JDBC 目录中的一个表或视图。</summary>
    internal sealed class JdbcObjectInfo
    {
        /// <summary>获取或设置架构名。</summary>
        public string Schema { get; set; }

        /// <summary>获取或设置对象名。</summary>
        public string Name { get; set; }

        /// <summary>获取或设置对象种类，table 或 view。</summary>
        public string Kind { get; set; }
    }

    /// <summary>把 JDBC 转接请求和响应编码为单行 JSON。</summary>
    internal static class JdbcProtocol
    {
        /// <summary>序列化一条请求。结果可能包含密码，不能写入日志。</summary>
        /// <param name="request">请求内容。</param>
        /// <returns>单行 JSON。</returns>
        public static string WriteRequest(JdbcBridgeRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            Dictionary<string, object> payload = new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["id"] = request.Id ?? string.Empty,
                ["op"] = request.Op ?? string.Empty,
                ["queryId"] = request.QueryId ?? string.Empty,
                ["url"] = request.Url ?? string.Empty,
                ["user"] = request.User ?? string.Empty,
                ["password"] = request.Password ?? string.Empty,
                ["driver"] = request.Driver ?? string.Empty,
                ["jars"] = (request.Jars ?? new List<string>()).Where(item => !string.IsNullOrWhiteSpace(item)).Select(item => item.Trim()).ToList(),
                ["sql"] = request.Sql ?? string.Empty,
                ["catalog"] = request.Catalog ?? string.Empty,
                ["timeoutSeconds"] = request.TimeoutSeconds
            };
            return JdbcJson.Serialize(payload);
        }

        /// <summary>解析转接进程写出的一帧响应。</summary>
        /// <param name="line">一行 JSON。</param>
        /// <returns>响应对象。</returns>
        /// <exception cref="FormatException">该行不是 JSON 对象。</exception>
        public static JdbcBridgeReply ReadReply(string line)
        {
            object parsed = JdbcJson.Parse(line);
            Dictionary<string, object> payload = parsed as Dictionary<string, object>;
            if (payload == null)
            {
                throw new FormatException("JDBC 响应必须是 JSON 对象。");
            }

            JdbcBridgeReply reply = new JdbcBridgeReply
            {
                Id = Text(payload, "id"),
                Type = Text(payload, "type"),
                Message = Text(payload, "message"),
                Columns = ReadColumns(payload),
                Rows = ReadRows(payload),
                Names = ReadNames(payload),
                Objects = ReadObjects(payload)
            };
            return reply;
        }

        /// <summary>按列类型把 JSON 单元格转换成 CLR 值。</summary>
        /// <param name="value">JSON 解析后的原始值。</param>
        /// <param name="clrType">列类型短名。</param>
        /// <returns>可供预览和 Excel 写入的值；空单元格返回 null。</returns>
        public static object ConvertCell(object value, string clrType)
        {
            if (value == null)
            {
                return null;
            }

            string type = clrType ?? "string";
            if (string.Equals(type, "bool", StringComparison.OrdinalIgnoreCase))
            {
                if (value is bool flag) return flag;
                return string.Equals(Convert.ToString(value, CultureInfo.InvariantCulture), "true", StringComparison.OrdinalIgnoreCase);
            }

            if (string.Equals(type, "int", StringComparison.OrdinalIgnoreCase))
            {
                long number = ToInt64(value);
                if (number >= int.MinValue && number <= int.MaxValue) return (int)number;
                return number;
            }

            if (string.Equals(type, "long", StringComparison.OrdinalIgnoreCase))
            {
                return ToInt64(value);
            }

            if (string.Equals(type, "double", StringComparison.OrdinalIgnoreCase))
            {
                if (value is JdbcNumber number) return double.Parse(number.Text, CultureInfo.InvariantCulture);
                if (value is double existing) return existing;
                return double.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
            }

            if (string.Equals(type, "decimal", StringComparison.OrdinalIgnoreCase))
            {
                string text = value is JdbcNumber number
                    ? number.Text
                    : Convert.ToString(value, CultureInfo.InvariantCulture);
                return decimal.Parse(text, CultureInfo.InvariantCulture);
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        /// <summary>把类型短名映射为结果列使用的 CLR 类型。</summary>
        /// <param name="clrType">转接进程给出的类型短名。</param>
        /// <returns>对应的 CLR 类型；未知时返回 <see cref="string"/>。</returns>
        public static Type ToClrType(string clrType)
        {
            if (string.Equals(clrType, "int", StringComparison.OrdinalIgnoreCase)) return typeof(int);
            if (string.Equals(clrType, "long", StringComparison.OrdinalIgnoreCase)) return typeof(long);
            if (string.Equals(clrType, "double", StringComparison.OrdinalIgnoreCase)) return typeof(double);
            if (string.Equals(clrType, "decimal", StringComparison.OrdinalIgnoreCase)) return typeof(decimal);
            if (string.Equals(clrType, "bool", StringComparison.OrdinalIgnoreCase)) return typeof(bool);
            return typeof(string);
        }

        private static long ToInt64(object value)
        {
            if (value is JdbcNumber number) return long.Parse(number.Text, CultureInfo.InvariantCulture);
            if (value is int integer) return integer;
            if (value is long whole) return whole;
            return long.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        }

        private static List<JdbcColumnInfo> ReadColumns(Dictionary<string, object> payload)
        {
            List<object> items = payload.TryGetValue("columns", out object value) ? value as List<object> : null;
            if (items == null) return new List<JdbcColumnInfo>();
            List<JdbcColumnInfo> columns = new List<JdbcColumnInfo>();
            foreach (object item in items)
            {
                Dictionary<string, object> column = item as Dictionary<string, object>;
                if (column == null) continue;
                columns.Add(new JdbcColumnInfo
                {
                    Name = Text(column, "name"),
                    Type = Text(column, "type")
                });
            }

            return columns;
        }

        private static List<object[]> ReadRows(Dictionary<string, object> payload)
        {
            List<object> items = payload.TryGetValue("rows", out object value) ? value as List<object> : null;
            if (items == null) return new List<object[]>();
            List<object[]> rows = new List<object[]>();
            foreach (object item in items)
            {
                List<object> cells = item as List<object> ?? new List<object>();
                rows.Add(cells.ToArray());
            }

            return rows;
        }

        private static List<string> ReadNames(Dictionary<string, object> payload)
        {
            List<object> items = payload.TryGetValue("names", out object value) ? value as List<object> : null;
            List<string> names = new List<string>();
            if (items == null) return names;
            foreach (object item in items)
            {
                string name = Convert.ToString(item, CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(name)) names.Add(name);
            }

            return names;
        }

        private static List<JdbcObjectInfo> ReadObjects(Dictionary<string, object> payload)
        {
            List<object> items = payload.TryGetValue("objects", out object value) ? value as List<object> : null;
            List<JdbcObjectInfo> objects = new List<JdbcObjectInfo>();
            if (items == null) return objects;
            foreach (object item in items)
            {
                Dictionary<string, object> row = item as Dictionary<string, object>;
                if (row == null) continue;
                objects.Add(new JdbcObjectInfo
                {
                    Schema = Text(row, "schema"),
                    Name = Text(row, "name"),
                    Kind = Text(row, "kind")
                });
            }

            return objects;
        }

        private static string Text(Dictionary<string, object> payload, string key)
        {
            if (!payload.TryGetValue(key, out object value) || value == null) return string.Empty;
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }
    }
}
