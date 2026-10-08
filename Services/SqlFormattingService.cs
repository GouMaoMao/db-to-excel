using System;
using System.Collections.Generic;
using System.Text;
using DB2Sheet.Contracts;

namespace DB2Sheet.Services
{
    /// <summary>以内部词法分析实现常见查询语句的结构化格式化。</summary>
    /// <remarks>服务不连接数据库且不修改 SQL 语义；它覆盖四类已支持数据库的常用查询语法，但不是完整方言解析器。格式化按列对齐子句内容，不使用固定缩进。词法标记只供格式化使用，不对外着色。</remarks>
    public sealed class SqlFormattingService : ISqlFormattingService
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "all", "and", "as", "asc", "between", "by", "case", "cast", "cross", "current", "desc", "distinct",
            "else", "end", "except", "exists", "false", "fetch", "first", "following", "for", "from", "full",
            "group", "having", "in", "inner", "intersect", "into", "is", "join", "last", "lateral", "left", "like",
            "limit", "not", "null", "nulls", "offset", "on", "or", "order", "outer", "over", "partition", "preceding",
            "recursive", "right", "rows", "select", "then", "true", "union", "using", "when", "where", "window", "with"
        };

        /// <summary>把 SQL 拆成不含空白的标记，供格式化区分字符串、注释和关键字。</summary>
        /// <param name="sql">待分析 SQL。空白会被跳过。</param>
        /// <returns>按出现顺序排列的标记。</returns>
        private static IReadOnlyList<SqlToken> Tokenize(string sql)
        {
            string value = sql ?? string.Empty;
            List<SqlToken> tokens = new List<SqlToken>();
            int index = 0;
            while (index < value.Length)
            {
                if (char.IsWhiteSpace(value[index]))
                {
                    index++;
                    continue;
                }

                int start = index;
                char current = value[index];
                char next = index + 1 < value.Length ? value[index + 1] : '\0';
                if (current == '-' && next == '-')
                {
                    index += 2;
                    while (index < value.Length && value[index] != '\r' && value[index] != '\n') index++;
                    Add(tokens, SqlTokenKind.Comment, value, start, index);
                    continue;
                }
                if (current == '/' && next == '*')
                {
                    index += 2;
                    while (index < value.Length && !(value[index - 1] == '*' && value[index] == '/')) index++;
                    if (index < value.Length) index++;
                    Add(tokens, SqlTokenKind.Comment, value, start, index);
                    continue;
                }
                if (current == '\'' || current == '"' || current == '`' || current == '[')
                {
                    char closing = current == '[' ? ']' : current;
                    bool isString = current == '\'';
                    index++;
                    while (index < value.Length)
                    {
                        if (value[index] == closing)
                        {
                            if (index + 1 < value.Length && value[index + 1] == closing)
                            {
                                index += 2;
                                continue;
                            }
                            index++;
                            break;
                        }
                        index++;
                    }
                    Add(tokens, isString ? SqlTokenKind.String : SqlTokenKind.QuotedIdentifier, value, start, index);
                    continue;
                }
                if (char.IsDigit(current))
                {
                    index++;
                    while (index < value.Length && (char.IsDigit(value[index]) || value[index] == '.')) index++;
                    Add(tokens, SqlTokenKind.Number, value, start, index);
                    continue;
                }
                if ((current == '@' || current == ':' || current == '$' || current == '?') &&
                    (next == '\0' || char.IsLetterOrDigit(next) || next == '_'))
                {
                    index++;
                    while (index < value.Length && (char.IsLetterOrDigit(value[index]) || value[index] == '_')) index++;
                    Add(tokens, SqlTokenKind.Parameter, value, start, index);
                    continue;
                }
                if (char.IsLetter(current) || current == '_')
                {
                    index++;
                    while (index < value.Length && (char.IsLetterOrDigit(value[index]) || value[index] == '_' || value[index] == '$')) index++;
                    string text = value.Substring(start, index - start);
                    tokens.Add(new SqlToken(Keywords.Contains(text) ? SqlTokenKind.Keyword : SqlTokenKind.Word, text));
                    continue;
                }

                index++;
                if (index < value.Length && IsPairOperator(current, value[index])) index++;
                Add(tokens, SqlTokenKind.Symbol, value, start, index);
            }
            return tokens.AsReadOnly();
        }

        /// <inheritdoc/>
        public string Format(string sql)
        {
            IReadOnlyList<SqlToken> tokens = Tokenize(sql);
            if (tokens.Count == 0) return string.Empty;
            return new Layout(tokens).Render().Trim();
        }

        private static void Add(ICollection<SqlToken> tokens, SqlTokenKind kind, string source, int start, int end)
        {
            tokens.Add(new SqlToken(kind, source.Substring(start, end - start)));
        }

        private static bool IsPairOperator(char first, char second)
        {
            return (first == '<' && (second == '=' || second == '>')) ||
                   (first == '>' && second == '=') ||
                   (first == '!' && second == '=') ||
                   (first == ':' && second == ':') ||
                   (first == '|' && second == '|');
        }

        /// <summary>一次遍历标记，按当前语句的对齐列输出 SQL。</summary>
        /// <remarks>每个 <see cref="Format"/> 调用使用独立实例。函数参数、值列表、窗口和普通括号在本地深度内单行输出，不进入子句换行。</remarks>
        private sealed class Layout
        {
            private readonly IReadOnlyList<SqlToken> _tokens;
            private readonly StringBuilder _text = new StringBuilder();
            private readonly List<Frame> _frames = new List<Frame>();
            private readonly Stack<CaseMark> _cases = new Stack<CaseMark>();
            private int _index;
            private int _column;
            private int _depth;
            private bool _lastWasKeyword;
            private bool _lastWasWord;

            /// <summary>用已经去掉空白的标记创建一次排版。</summary>
            /// <param name="tokens">词法标记。</param>
            public Layout(IReadOnlyList<SqlToken> tokens)
            {
                _tokens = tokens;
                _frames.Add(new Frame());
            }

            /// <summary>输出对齐后的 SQL。</summary>
            /// <returns>未修剪首尾空白的排版结果。</returns>
            public string Render()
            {
                while (_index < _tokens.Count)
                {
                    SqlToken token = _tokens[_index];
                    if (token.Kind == SqlTokenKind.Comment)
                    {
                        WriteComment(token);
                        _index++;
                        continue;
                    }
                    if (TryHandleCase()) continue;
                    if (TryStartClause()) continue;
                    if (IsSymbol("("))
                    {
                        HandleOpenParen();
                        continue;
                    }
                    if (IsSymbol(")"))
                    {
                        HandleCloseParen();
                        continue;
                    }
                    if (IsSymbol(","))
                    {
                        HandleComma();
                        continue;
                    }
                    if (IsSymbol(";"))
                    {
                        HandleSemicolon();
                        continue;
                    }
                    if (TryHandleAndOr()) continue;
                    if (TryHandleOn()) continue;
                    if (IsSelectModifier(token))
                    {
                        AppendDisplay(token, false);
                        _index++;
                        continue;
                    }

                    AppendDisplay(token, true);
                    _index++;
                }
                return _text.ToString();
            }

            private bool TryStartClause()
            {
                if (InCurrentCase() || _depth != Top.Depth) return false;
                string text;
                ClauseKind kind;
                int count;
                if (!TryReadClause(out text, out kind, out count)) return false;
                StartClause(text, kind);
                _index += count;
                return true;
            }

            private bool TryReadClause(out string text, out ClauseKind kind, out int count)
            {
                text = null;
                kind = ClauseKind.Plain;
                count = 0;
                if (!IsKw(0)) return false;
                if (IsKw(0, "group") && IsKw(1, "by"))
                {
                    text = "group by";
                    kind = ClauseKind.List;
                    count = 2;
                    return true;
                }
                if (IsKw(0, "order") && IsKw(1, "by"))
                {
                    text = "order by";
                    kind = ClauseKind.List;
                    count = 2;
                    return true;
                }

                string join;
                int joinCount;
                if (TryReadJoin(out join, out joinCount))
                {
                    text = join;
                    kind = ClauseKind.Join;
                    count = joinCount;
                    return true;
                }
                if (IsKw(0, "union") || IsKw(0, "except") || IsKw(0, "intersect"))
                {
                    text = LowerAt(0);
                    count = 1;
                    if (IsKw(1, "all"))
                    {
                        text = text + " all";
                        count = 2;
                    }
                    kind = ClauseKind.Plain;
                    return true;
                }

                string word = LowerAt(0);
                if (word == "select")
                {
                    text = word;
                    kind = ClauseKind.SelectList;
                    count = 1;
                    return true;
                }
                if (word == "with")
                {
                    text = word;
                    kind = ClauseKind.List;
                    count = 1;
                    return true;
                }
                if (word == "where" || word == "having")
                {
                    text = word;
                    kind = ClauseKind.Boolean;
                    count = 1;
                    return true;
                }
                if (word == "from" || word == "limit" || word == "offset" || word == "fetch" || word == "window")
                {
                    text = word;
                    kind = ClauseKind.Plain;
                    count = 1;
                    return true;
                }
                return false;
            }

            private bool TryReadJoin(out string text, out int count)
            {
                text = null;
                count = 0;
                if (IsKw(0, "join"))
                {
                    text = "join";
                    count = 1;
                    return true;
                }
                if (!IsKw(0, "left") && !IsKw(0, "right") && !IsKw(0, "full") &&
                    !IsKw(0, "inner") && !IsKw(0, "cross") && !IsKw(0, "outer"))
                {
                    return false;
                }

                int end = 1;
                if (IsKw(1, "outer")) end = 2;
                if (!IsKw(end, "join")) return false;
                StringBuilder phrase = new StringBuilder();
                for (int offset = 0; offset <= end; offset++)
                {
                    if (offset > 0) phrase.Append(' ');
                    phrase.Append(LowerAt(offset));
                }
                text = phrase.ToString();
                count = end + 1;
                return true;
            }

            private void StartClause(string text, ClauseKind kind)
            {
                bool stick = Top.AwaitingStart && (text == "select" || text == "with") && LastChar() == '(';
                Top.AwaitingStart = false;
                if (!stick) NewLine(Top.BaseColumn);
                AppendRaw(text);
                Top.Clause = kind;
                Top.ContentStarted = false;
            }

            private bool TryHandleCase()
            {
                SqlToken token = _tokens[_index];
                if (token.Kind != SqlTokenKind.Keyword) return false;
                string word = token.Text.ToLowerInvariant();
                if (word == "case")
                {
                    BeginCase();
                    _index++;
                    return true;
                }
                if (!InCurrentCase()) return false;

                CaseMark mark = _cases.Peek();
                if (word == "when" && mark.Phase != CasePhase.Condition)
                {
                    NewLine(mark.Column + 2);
                    AppendRaw("when");
                    mark.Phase = CasePhase.Condition;
                    _index++;
                    return true;
                }
                if (word == "else")
                {
                    NewLine(mark.Column + 2);
                    AppendRaw("else");
                    mark.Phase = CasePhase.Result;
                    _index++;
                    return true;
                }
                if (word == "then" && mark.Phase == CasePhase.Condition)
                {
                    AppendDisplay(token, false);
                    mark.Phase = CasePhase.Result;
                    _index++;
                    return true;
                }
                if (word == "end")
                {
                    NewLine(mark.Column);
                    AppendRaw("end");
                    _cases.Pop();
                    _index++;
                    return true;
                }
                return false;
            }

            private void BeginCase()
            {
                if (NeedsSpaceBefore("case"))
                {
                    _text.Append(' ');
                    _column++;
                }
                if (!Top.ContentStarted)
                {
                    Top.ContentColumn = _column;
                    Top.ContentStarted = true;
                }
                int column = _column;
                AppendRaw("case");
                _cases.Push(new CaseMark
                {
                    Column = column,
                    Depth = _depth,
                    FrameCount = _frames.Count,
                    Phase = CasePhase.Operand
                });
            }

            private void HandleOpenParen()
            {
                if (!_lastWasWord)
                {
                    int next = NextSignificant(_index + 1);
                    if (next >= 0 && (IsKeywordAt(next, "select") || IsKeywordAt(next, "with")))
                    {
                        AppendDisplay(_tokens[_index], true);
                        _index++;
                        _depth++;
                        _frames.Add(new Frame
                        {
                            BaseColumn = _column,
                            Depth = _depth,
                            AwaitingStart = true
                        });
                        return;
                    }
                    if (next >= 0 && IsKeywordAt(next, "case"))
                    {
                        AppendDisplay(_tokens[_index], true);
                        _index++;
                        _depth++;
                        return;
                    }
                }

                WriteInlineParen();
            }

            /// <summary>把当前括号直到匹配的右括号写成一行。调用时当前标记是左括号。</summary>
            private void WriteInlineParen()
            {
                AppendDisplay(_tokens[_index], true);
                _index++;
                int depth = 1;
                while (_index < _tokens.Count && depth > 0)
                {
                    SqlToken token = _tokens[_index];
                    if (token.Kind == SqlTokenKind.Comment)
                    {
                        WriteComment(token);
                        _index++;
                        continue;
                    }
                    if (token.Text == "(") depth++;
                    else if (token.Text == ")") depth--;
                    if (token.Text == "," && depth > 0)
                    {
                        _text.Append(", ");
                        _column += 2;
                        _lastWasKeyword = false;
                        _lastWasWord = false;
                    }
                    else
                    {
                        AppendDisplay(token, false);
                    }
                    _index++;
                }
            }

            private void HandleCloseParen()
            {
                if (_frames.Count > 1 && Top.Depth == _depth)
                    _frames.RemoveAt(_frames.Count - 1);
                if (_depth > 0) _depth--;
                AppendPlain(")");
                _index++;
            }

            private void HandleComma()
            {
                bool list = !InCurrentCase() && _depth == Top.Depth &&
                            (Top.Clause == ClauseKind.SelectList || Top.Clause == ClauseKind.List);
                AppendPlain(",");
                _index++;
                if (list) NewLine(Top.ContentColumn);
                else
                {
                    _text.Append(' ');
                    _column++;
                }
            }

            private void HandleSemicolon()
            {
                AppendPlain(";");
                _index++;
                _depth = 0;
                _frames.Clear();
                _frames.Add(new Frame());
                _cases.Clear();
                if (_index < _tokens.Count) NewLine(0);
            }

            private bool TryHandleAndOr()
            {
                if (InCurrentCase() || _depth != Top.Depth || Top.Clause != ClauseKind.Boolean) return false;
                SqlToken token = _tokens[_index];
                if (token.Kind != SqlTokenKind.Keyword) return false;
                if (!token.Text.Equals("and", StringComparison.OrdinalIgnoreCase) &&
                    !token.Text.Equals("or", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
                NewLine(Top.ContentColumn);
                AppendRaw(token.Text.ToLowerInvariant());
                _index++;
                return true;
            }

            private bool TryHandleOn()
            {
                if (InCurrentCase() || _depth != Top.Depth || Top.Clause != ClauseKind.Join) return false;
                SqlToken token = _tokens[_index];
                if (token.Kind != SqlTokenKind.Keyword || !token.Text.Equals("on", StringComparison.OrdinalIgnoreCase))
                    return false;
                AppendDisplay(token, false);
                Top.Clause = ClauseKind.Boolean;
                Top.ContentStarted = false;
                _index++;
                return true;
            }

            private void WriteComment(SqlToken token)
            {
                if (!AtLineStart() || _column != Top.BaseColumn)
                    NewLine(Top.BaseColumn);
                _text.Append(token.Text);
                int lastBreak = token.Text.LastIndexOf('\n');
                _column = lastBreak >= 0 ? token.Text.Length - lastBreak - 1 : _column + token.Text.Length;
                if (_index + 1 < _tokens.Count && (_text.Length == 0 || _text[_text.Length - 1] != '\n'))
                {
                    _text.Append('\n');
                    _column = 0;
                }
            }

            private void AppendDisplay(SqlToken token, bool noteContent)
            {
                string text = token.Kind == SqlTokenKind.Keyword ? token.Text.ToLowerInvariant() : token.Text;
                if (NeedsSpaceBefore(text))
                {
                    _text.Append(' ');
                    _column++;
                }
                if (noteContent && !Top.ContentStarted)
                {
                    Top.ContentColumn = _column;
                    Top.ContentStarted = true;
                }
                _text.Append(text);
                _column += text.Length;
                _lastWasKeyword = token.Kind == SqlTokenKind.Keyword;
                _lastWasWord = token.Kind == SqlTokenKind.Word;
            }

            private void AppendRaw(string text)
            {
                _text.Append(text);
                _column += text.Length;
                _lastWasKeyword = true;
                _lastWasWord = false;
            }

            private void AppendPlain(string text)
            {
                _text.Append(text);
                _column += text.Length;
                _lastWasKeyword = false;
                _lastWasWord = false;
            }

            private void NewLine(int column)
            {
                if (column < 0) column = 0;
                if (_text.Length > 0 && _text[_text.Length - 1] != '\n')
                    _text.Append('\n');
                if (column > 0)
                    _text.Append(new string(' ', column));
                _column = column;
            }

            private bool NeedsSpaceBefore(string text)
            {
                if (_text.Length == 0) return false;
                char last = _text[_text.Length - 1];
                if (last == '\n' || last == ' ' || last == '(' || last == '.') return false;
                if (text == "." || text == "," || text == ")" || text == ";") return false;
                if (text == "(") return _lastWasKeyword;
                return true;
            }

            private bool InCurrentCase()
            {
                if (_cases.Count == 0) return false;
                CaseMark mark = _cases.Peek();
                return mark.FrameCount == _frames.Count && mark.Depth == _depth;
            }

            private bool IsSelectModifier(SqlToken token)
            {
                if (InCurrentCase() || Top.ContentStarted || Top.Clause != ClauseKind.SelectList) return false;
                if (token.Kind != SqlTokenKind.Keyword) return false;
                return token.Text.Equals("distinct", StringComparison.OrdinalIgnoreCase) ||
                       token.Text.Equals("all", StringComparison.OrdinalIgnoreCase);
            }

            private bool IsSymbol(string text)
            {
                SqlToken token = _tokens[_index];
                return token.Kind == SqlTokenKind.Symbol && token.Text == text;
            }

            private bool IsKw(int offset)
            {
                int index = _index + offset;
                return index >= 0 && index < _tokens.Count && _tokens[index].Kind == SqlTokenKind.Keyword;
            }

            private bool IsKw(int offset, string word)
            {
                int index = _index + offset;
                return index >= 0 && index < _tokens.Count &&
                       _tokens[index].Kind == SqlTokenKind.Keyword &&
                       _tokens[index].Text.Equals(word, StringComparison.OrdinalIgnoreCase);
            }

            private bool IsKeywordAt(int index, string word)
            {
                return _tokens[index].Kind == SqlTokenKind.Keyword &&
                       _tokens[index].Text.Equals(word, StringComparison.OrdinalIgnoreCase);
            }

            private string LowerAt(int offset)
            {
                return _tokens[_index + offset].Text.ToLowerInvariant();
            }

            private int NextSignificant(int start)
            {
                for (int index = start; index < _tokens.Count; index++)
                {
                    if (_tokens[index].Kind != SqlTokenKind.Comment) return index;
                }
                return -1;
            }

            private bool AtLineStart()
            {
                return _text.Length == 0 || _text[_text.Length - 1] == '\n';
            }

            private char LastChar()
            {
                return _text.Length == 0 ? '\0' : _text[_text.Length - 1];
            }

            private Frame Top
            {
                get { return _frames[_frames.Count - 1]; }
            }

            /// <summary>记录一层 select 或 with 语句的对齐位置。</summary>
            private sealed class Frame
            {
                /// <summary>子句关键字所在列。根语句为 0，子查询为左括号后的列。</summary>
                public int BaseColumn;

                /// <summary>这一层语句对应的括号深度。根语句为 0。</summary>
                public int Depth;

                /// <summary>当前子句如何换行。</summary>
                public ClauseKind Clause;

                /// <summary>列表项或布尔条件的起始列。</summary>
                public int ContentColumn;

                /// <summary>是否已经写出该子句的第一项。为假时下一项确定对齐列。</summary>
                public bool ContentStarted;

                /// <summary>子查询左括号已经写出，内部 select 或 with 还要贴在括号后。</summary>
                public bool AwaitingStart;
            }

            /// <summary>记录尚未结束的 case，便于 when 和 end 对齐到 case 本身。</summary>
            private sealed class CaseMark
            {
                /// <summary>case 关键字的起始列。</summary>
                public int Column;

                /// <summary>case 所在的括号深度。</summary>
                public int Depth;

                /// <summary>打开 case 时的语句帧数量。进入子查询后外层 case 暂停换行。</summary>
                public int FrameCount;

                /// <summary>当前处于操作数、when 条件还是 then 结果。</summary>
                public CasePhase Phase;
            }

            /// <summary>决定逗号和 AND/OR 要不要换行。</summary>
            private enum ClauseKind
            {
                /// <summary>尚未进入子句。</summary>
                None,

                /// <summary>select 列表，顶层逗号换行。</summary>
                SelectList,

                /// <summary>group by、order by 或 with，顶层逗号换行。</summary>
                List,

                /// <summary>where、having 或 join 的 on，顶层 AND/OR 换行。</summary>
                Boolean,

                /// <summary>join 行。遇到 on 后改为布尔换行。</summary>
                Join,

                /// <summary>from、limit 等保持在同一行的子句。</summary>
                Plain
            }

            /// <summary>case 表达式写到了哪一段。</summary>
            private enum CasePhase
            {
                /// <summary>case 与第一个 when 之间的判别式。</summary>
                Operand,

                /// <summary>when 与 then 之间的条件。</summary>
                Condition,

                /// <summary>then 或 else 之后的结果。</summary>
                Result
            }
        }

        /// <summary>区分格式化所用的词法标记。字符串、注释和关键字必须分开，避免改写字面量。</summary>
        private enum SqlTokenKind
        {
            /// <summary>SQL 保留字或常用子句关键字。</summary>
            Keyword,
            /// <summary>普通名称。</summary>
            Word,
            /// <summary>单引号字符串。</summary>
            String,
            /// <summary>单行或块注释。</summary>
            Comment,
            /// <summary>整数或小数。</summary>
            Number,
            /// <summary>以 @、:、$ 或 ? 表示的参数。</summary>
            Parameter,
            /// <summary>由双引号、反引号或方括号包围的标识符。</summary>
            QuotedIdentifier,
            /// <summary>运算符、标点或括号。</summary>
            Symbol
        }

        /// <summary>格式化使用的词法标记，只保留种类和原文。</summary>
        private sealed class SqlToken
        {
            /// <summary>创建格式化用的词法标记。</summary>
            /// <param name="kind">标记类型。</param>
            /// <param name="text">标记原始文本。</param>
            public SqlToken(SqlTokenKind kind, string text)
            {
                Kind = kind;
                Text = text ?? string.Empty;
            }

            /// <summary>获取标记类型。</summary>
            public SqlTokenKind Kind { get; }

            /// <summary>获取标记原始文本。</summary>
            public string Text { get; }
        }
    }
}
