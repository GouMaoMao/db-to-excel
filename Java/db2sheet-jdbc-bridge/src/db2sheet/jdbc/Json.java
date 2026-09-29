package db2sheet.jdbc;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 转接协议使用的最小 JSON 读写器，避免依赖第三方库。
 * 只覆盖对象、数组、字符串、数字、布尔和 null。
 */
final class Json {
    private Json() {
    }

    static String write(Object value) {
        StringBuilder builder = new StringBuilder();
        writeValue(builder, value);
        return builder.toString();
    }

    static Map<String, Object> parseObject(String text) {
        Object value = new Parser(text).parseValue();
        if (!(value instanceof Map)) {
            throw new IllegalArgumentException("JSON 根节点必须是对象。");
        }
        @SuppressWarnings("unchecked")
        Map<String, Object> object = (Map<String, Object>) value;
        return object;
    }

    static String text(Map<String, Object> object, String key) {
        Object value = object.get(key);
        return value == null ? "" : String.valueOf(value);
    }

    static int integer(Map<String, Object> object, String key, int fallback) {
        Object value = object.get(key);
        if (value instanceof Number) {
            return ((Number) value).intValue();
        }
        if (value == null) {
            return fallback;
        }
        try {
            return Integer.parseInt(String.valueOf(value));
        } catch (NumberFormatException exception) {
            return fallback;
        }
    }

    private static void writeValue(StringBuilder builder, Object value) {
        if (value == null) {
            builder.append("null");
        } else if (value instanceof String) {
            writeString(builder, (String) value);
        } else if (value instanceof Boolean || value instanceof Integer || value instanceof Long) {
            builder.append(value.toString());
        } else if (value instanceof Double || value instanceof Float) {
            double number = ((Number) value).doubleValue();
            if (Double.isNaN(number) || Double.isInfinite(number)) {
                writeString(builder, Double.toString(number));
            } else {
                builder.append(Double.toString(number));
            }
        } else if (value instanceof Number) {
            builder.append(value.toString());
        } else if (value instanceof Map) {
            builder.append('{');
            boolean first = true;
            for (Map.Entry<?, ?> entry : ((Map<?, ?>) value).entrySet()) {
                if (!first) {
                    builder.append(',');
                }
                first = false;
                writeString(builder, String.valueOf(entry.getKey()));
                builder.append(':');
                writeValue(builder, entry.getValue());
            }
            builder.append('}');
        } else if (value instanceof Iterable) {
            builder.append('[');
            boolean first = true;
            for (Object item : (Iterable<?>) value) {
                if (!first) {
                    builder.append(',');
                }
                first = false;
                writeValue(builder, item);
            }
            builder.append(']');
        } else {
            writeString(builder, String.valueOf(value));
        }
    }

    private static void writeString(StringBuilder builder, String value) {
        builder.append('"');
        for (int index = 0; index < value.length(); index++) {
            char current = value.charAt(index);
            switch (current) {
                case '"':
                    builder.append("\\\"");
                    break;
                case '\\':
                    builder.append("\\\\");
                    break;
                case '\b':
                    builder.append("\\b");
                    break;
                case '\f':
                    builder.append("\\f");
                    break;
                case '\n':
                    builder.append("\\n");
                    break;
                case '\r':
                    builder.append("\\r");
                    break;
                case '\t':
                    builder.append("\\t");
                    break;
                default:
                    if (current < 0x20) {
                        builder.append(String.format("\\u%04x", (int) current));
                    } else {
                        builder.append(current);
                    }
                    break;
            }
        }
        builder.append('"');
    }

    private static final class Parser {
        private final String text;
        private int index;

        private Parser(String text) {
            this.text = text == null ? "" : text;
        }

        private Object parseValue() {
            skipSpace();
            if (index >= text.length()) {
                throw new IllegalArgumentException("JSON 内容为空。");
            }
            char current = text.charAt(index);
            if (current == '{') {
                return parseObject();
            }
            if (current == '[') {
                return parseArray();
            }
            if (current == '"') {
                return parseString();
            }
            if (current == 't' || current == 'f') {
                return parseBoolean();
            }
            if (current == 'n') {
                return parseNull();
            }
            return parseNumber();
        }

        private Map<String, Object> parseObject() {
            expect('{');
            Map<String, Object> object = new LinkedHashMap<String, Object>();
            skipSpace();
            if (peek('}')) {
                index++;
                return object;
            }
            while (index < text.length()) {
                skipSpace();
                String key = parseString();
                skipSpace();
                expect(':');
                object.put(key, parseValue());
                skipSpace();
                if (peek('}')) {
                    index++;
                    return object;
                }
                expect(',');
            }
            throw new IllegalArgumentException("JSON 对象未闭合。");
        }

        private List<Object> parseArray() {
            expect('[');
            List<Object> array = new ArrayList<Object>();
            skipSpace();
            if (peek(']')) {
                index++;
                return array;
            }
            while (index < text.length()) {
                array.add(parseValue());
                skipSpace();
                if (peek(']')) {
                    index++;
                    return array;
                }
                expect(',');
            }
            throw new IllegalArgumentException("JSON 数组未闭合。");
        }

        private String parseString() {
            expect('"');
            StringBuilder builder = new StringBuilder();
            while (index < text.length()) {
                char current = text.charAt(index++);
                if (current == '"') {
                    return builder.toString();
                }
                if (current == '\\') {
                    if (index >= text.length()) {
                        throw new IllegalArgumentException("JSON 字符串转义未完成。");
                    }
                    char escaped = text.charAt(index++);
                    switch (escaped) {
                        case '"':
                        case '\\':
                        case '/':
                            builder.append(escaped);
                            break;
                        case 'b':
                            builder.append('\b');
                            break;
                        case 'f':
                            builder.append('\f');
                            break;
                        case 'n':
                            builder.append('\n');
                            break;
                        case 'r':
                            builder.append('\r');
                            break;
                        case 't':
                            builder.append('\t');
                            break;
                        case 'u':
                            if (index + 4 > text.length()) {
                                throw new IllegalArgumentException("JSON Unicode 转义未完成。");
                            }
                            int code = Integer.parseInt(text.substring(index, index + 4), 16);
                            builder.append((char) code);
                            index += 4;
                            break;
                        default:
                            throw new IllegalArgumentException("JSON 字符串包含未知转义。");
                    }
                } else if (current < 0x20) {
                    throw new IllegalArgumentException("JSON 字符串包含未转义的控制字符。");
                } else {
                    builder.append(current);
                }
            }
            throw new IllegalArgumentException("JSON 字符串未闭合。");
        }

        private Boolean parseBoolean() {
            if (text.startsWith("true", index)) {
                index += 4;
                return Boolean.TRUE;
            }
            if (text.startsWith("false", index)) {
                index += 5;
                return Boolean.FALSE;
            }
            throw new IllegalArgumentException("JSON 布尔值无效。");
        }

        private Object parseNull() {
            if (text.startsWith("null", index)) {
                index += 4;
                return null;
            }
            throw new IllegalArgumentException("JSON null 无效。");
        }

        private Number parseNumber() {
            int start = index;
            if (peek('-')) {
                index++;
            }
            if (index >= text.length() || !Character.isDigit(text.charAt(index))) {
                throw new IllegalArgumentException("JSON 数字无效。");
            }
            while (index < text.length() && Character.isDigit(text.charAt(index))) {
                index++;
            }
            boolean fractional = false;
            if (peek('.')) {
                fractional = true;
                index++;
                while (index < text.length() && Character.isDigit(text.charAt(index))) {
                    index++;
                }
            }
            if (peek('e') || peek('E')) {
                fractional = true;
                index++;
                if (peek('+') || peek('-')) {
                    index++;
                }
                while (index < text.length() && Character.isDigit(text.charAt(index))) {
                    index++;
                }
            }
            String token = text.substring(start, index);
            if (!fractional) {
                try {
                    return Long.valueOf(token);
                } catch (NumberFormatException exception) {
                    return Double.valueOf(token);
                }
            }
            return Double.valueOf(token);
        }

        private void skipSpace() {
            while (index < text.length() && Character.isWhitespace(text.charAt(index))) {
                index++;
            }
        }

        private boolean peek(char expected) {
            return index < text.length() && text.charAt(index) == expected;
        }

        private void expect(char expected) {
            skipSpace();
            if (index >= text.length() || text.charAt(index) != expected) {
                throw new IllegalArgumentException("JSON 缺少字符 " + expected + "。");
            }
            index++;
        }
    }
}
