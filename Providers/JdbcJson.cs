using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace DB2Sheet.Providers
{
    /// <summary>为 JDBC 转接协议提供不依赖第三方库的 JSON 读写。</summary>
    /// <remarks>只支持对象、数组、字符串、数字、布尔和 null。数字保留原始文本，便于按列类型转换。</remarks>
    internal static class JdbcJson
    {
        /// <summary>把对象、数组、字符串、数字或布尔值写成一行 JSON。</summary>
        /// <param name="value">要序列化的值。</param>
        /// <returns>不含换行的 JSON 文本。</returns>
        public static string Serialize(object value)
        {
            StringBuilder builder = new StringBuilder();
            WriteValue(builder, value);
            return builder.ToString();
        }

        /// <summary>解析一行 JSON。</summary>
        /// <param name="text">JSON 文本。</param>
        /// <returns>字典、列表、字符串、<see cref="JdbcNumber"/>、布尔或 null。</returns>
        /// <exception cref="FormatException">JSON 结构无效。</exception>
        public static object Parse(string text)
        {
            Parser parser = new Parser(text ?? string.Empty);
            object value = parser.ParseValue();
            parser.SkipSpace();
            if (!parser.End)
            {
                throw new FormatException("JSON 末尾存在多余内容。");
            }

            return value;
        }

        private static void WriteValue(StringBuilder builder, object value)
        {
            if (value == null)
            {
                builder.Append("null");
                return;
            }

            if (value is string text)
            {
                WriteString(builder, text);
                return;
            }

            if (value is bool flag)
            {
                builder.Append(flag ? "true" : "false");
                return;
            }

            if (value is JdbcNumber number)
            {
                builder.Append(number.Text);
                return;
            }

            if (value is int || value is long || value is short || value is byte)
            {
                builder.Append(Convert.ToString(value, CultureInfo.InvariantCulture));
                return;
            }

            if (value is float || value is double || value is decimal)
            {
                builder.Append(Convert.ToString(value, CultureInfo.InvariantCulture));
                return;
            }

            if (value is IDictionary dictionary)
            {
                builder.Append('{');
                bool first = true;
                foreach (DictionaryEntry entry in dictionary)
                {
                    if (!first)
                    {
                        builder.Append(',');
                    }

                    first = false;
                    WriteString(builder, Convert.ToString(entry.Key, CultureInfo.InvariantCulture) ?? string.Empty);
                    builder.Append(':');
                    WriteValue(builder, entry.Value);
                }

                builder.Append('}');
                return;
            }

            if (value is IEnumerable sequence)
            {
                builder.Append('[');
                bool first = true;
                foreach (object item in sequence)
                {
                    if (!first)
                    {
                        builder.Append(',');
                    }

                    first = false;
                    WriteValue(builder, item);
                }

                builder.Append(']');
                return;
            }

            WriteString(builder, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
        }

        private static void WriteString(StringBuilder builder, string value)
        {
            builder.Append('"');
            foreach (char current in value)
            {
                switch (current)
                {
                    case '"':
                        builder.Append("\\\"");
                        break;
                    case '\\':
                        builder.Append("\\\\");
                        break;
                    case '\b':
                        builder.Append("\\b");
                        break;
                    case '\f':
                        builder.Append("\\f");
                        break;
                    case '\n':
                        builder.Append("\\n");
                        break;
                    case '\r':
                        builder.Append("\\r");
                        break;
                    case '\t':
                        builder.Append("\\t");
                        break;
                    default:
                        if (current < 0x20)
                        {
                            builder.Append("\\u");
                            builder.Append(((int)current).ToString("x4", CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            builder.Append(current);
                        }

                        break;
                }
            }

            builder.Append('"');
        }

        private sealed class Parser
        {
            private readonly string _text;
            private int _index;

            public Parser(string text)
            {
                _text = text;
            }

            public bool End => _index >= _text.Length;

            public object ParseValue()
            {
                SkipSpace();
                if (End)
                {
                    throw new FormatException("JSON 内容为空。");
                }

                char current = _text[_index];
                if (current == '{') return ParseObject();
                if (current == '[') return ParseArray();
                if (current == '"') return ParseString();
                if (current == 't' || current == 'f') return ParseBoolean();
                if (current == 'n') return ParseNull();
                return ParseNumber();
            }

            public void SkipSpace()
            {
                while (!End && char.IsWhiteSpace(_text[_index]))
                {
                    _index++;
                }
            }

            private Dictionary<string, object> ParseObject()
            {
                Expect('{');
                Dictionary<string, object> result = new Dictionary<string, object>(StringComparer.Ordinal);
                SkipSpace();
                if (Peek('}'))
                {
                    _index++;
                    return result;
                }

                while (!End)
                {
                    SkipSpace();
                    string key = ParseString();
                    SkipSpace();
                    Expect(':');
                    result[key] = ParseValue();
                    SkipSpace();
                    if (Peek('}'))
                    {
                        _index++;
                        return result;
                    }

                    Expect(',');
                }

                throw new FormatException("JSON 对象未闭合。");
            }

            private List<object> ParseArray()
            {
                Expect('[');
                List<object> result = new List<object>();
                SkipSpace();
                if (Peek(']'))
                {
                    _index++;
                    return result;
                }

                while (!End)
                {
                    result.Add(ParseValue());
                    SkipSpace();
                    if (Peek(']'))
                    {
                        _index++;
                        return result;
                    }

                    Expect(',');
                }

                throw new FormatException("JSON 数组未闭合。");
            }

            private string ParseString()
            {
                Expect('"');
                StringBuilder builder = new StringBuilder();
                while (!End)
                {
                    char current = _text[_index++];
                    if (current == '"')
                    {
                        return builder.ToString();
                    }

                    if (current == '\\')
                    {
                        if (End)
                        {
                            throw new FormatException("JSON 字符串转义未完成。");
                        }

                        char escaped = _text[_index++];
                        switch (escaped)
                        {
                            case '"':
                            case '\\':
                            case '/':
                                builder.Append(escaped);
                                break;
                            case 'b':
                                builder.Append('\b');
                                break;
                            case 'f':
                                builder.Append('\f');
                                break;
                            case 'n':
                                builder.Append('\n');
                                break;
                            case 'r':
                                builder.Append('\r');
                                break;
                            case 't':
                                builder.Append('\t');
                                break;
                            case 'u':
                                if (_index + 4 > _text.Length)
                                {
                                    throw new FormatException("JSON Unicode 转义未完成。");
                                }

                                int code = int.Parse(_text.Substring(_index, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                                builder.Append((char)code);
                                _index += 4;
                                break;
                            default:
                                throw new FormatException("JSON 字符串包含未知转义。");
                        }
                    }
                    else if (current < 0x20)
                    {
                        throw new FormatException("JSON 字符串包含未转义的控制字符。");
                    }
                    else
                    {
                        builder.Append(current);
                    }
                }

                throw new FormatException("JSON 字符串未闭合。");
            }

            private bool ParseBoolean()
            {
                if (_text.IndexOf("true", _index, StringComparison.Ordinal) == _index)
                {
                    _index += 4;
                    return true;
                }

                if (_text.IndexOf("false", _index, StringComparison.Ordinal) == _index)
                {
                    _index += 5;
                    return false;
                }

                throw new FormatException("JSON 布尔值无效。");
            }

            private object ParseNull()
            {
                if (_text.IndexOf("null", _index, StringComparison.Ordinal) == _index)
                {
                    _index += 4;
                    return null;
                }

                throw new FormatException("JSON null 无效。");
            }

            private JdbcNumber ParseNumber()
            {
                int start = _index;
                if (Peek('-'))
                {
                    _index++;
                }

                if (End || !char.IsDigit(_text[_index]))
                {
                    throw new FormatException("JSON 数字无效。");
                }

                while (!End && char.IsDigit(_text[_index]))
                {
                    _index++;
                }

                bool fractional = false;
                if (Peek('.'))
                {
                    fractional = true;
                    _index++;
                    while (!End && char.IsDigit(_text[_index]))
                    {
                        _index++;
                    }
                }

                if (Peek('e') || Peek('E'))
                {
                    fractional = true;
                    _index++;
                    if (Peek('+') || Peek('-'))
                    {
                        _index++;
                    }

                    while (!End && char.IsDigit(_text[_index]))
                    {
                        _index++;
                    }
                }

                return new JdbcNumber(_text.Substring(start, _index - start), !fractional);
            }

            private bool Peek(char expected)
            {
                return !End && _text[_index] == expected;
            }

            private void Expect(char expected)
            {
                SkipSpace();
                if (End || _text[_index] != expected)
                {
                    throw new FormatException("JSON 缺少字符 " + expected + "。");
                }

                _index++;
            }
        }
    }

    /// <summary>保留 JSON 数字的原始文本，避免在知道列类型之前丢失精度。</summary>
    internal sealed class JdbcNumber
    {
        /// <summary>创建 JSON 数字。</summary>
        /// <param name="text">原始数字文本。</param>
        /// <param name="isInteger">是否不含小数点和指数。</param>
        public JdbcNumber(string text, bool isInteger)
        {
            Text = text ?? string.Empty;
            IsInteger = isInteger;
        }

        /// <summary>获取原始数字文本。</summary>
        public string Text { get; }

        /// <summary>获取该数字是否为整数形式。</summary>
        public bool IsInteger { get; }
    }
}
