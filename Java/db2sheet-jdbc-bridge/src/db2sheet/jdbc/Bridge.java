package db2sheet.jdbc;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileDescriptor;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.Driver;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Types;
import java.util.ArrayList;
import java.util.Base64;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * 在独立 Java 进程中加载用户提供的 JDBC 驱动，并通过标准输入输出与插件交换查询结果。
 * 进程不包含任何厂商驱动；每次请求携带 jar 路径列表。
 */
public final class Bridge {
    private static final int RowBatchSize = 200;
    private static final String[] TableTypes = new String[] { "TABLE", "VIEW" };

    private final Object outputLock = new Object();
    private final Map<String, Session> sessions = new ConcurrentHashMap<String, Session>();
    private final Map<String, Driver> drivers = new ConcurrentHashMap<String, Driver>();
    private final Map<String, ClassLoader> loadersBySignature = new ConcurrentHashMap<String, ClassLoader>();
    private final BufferedWriter output;
    private final ExecutorService workers;

    /**
     * 启动转接循环。
     *
     * @param args 忽略；驱动 jar 由每条请求提供。
     */
    public static void main(String[] args) throws Exception {
        new Bridge().run();
    }

    private Bridge() throws Exception {
        this.output = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(FileDescriptor.out), StandardCharsets.UTF_8));
        this.workers = Executors.newCachedThreadPool(new java.util.concurrent.ThreadFactory() {
            @Override
            public Thread newThread(Runnable runnable) {
                Thread thread = new Thread(runnable, "db2sheet-jdbc");
                thread.setDaemon(true);
                return thread;
            }
        });
    }

    private void run() throws Exception {
        BufferedReader input = new BufferedReader(new InputStreamReader(new FileInputStream(FileDescriptor.in), StandardCharsets.UTF_8));
        String line;
        while ((line = input.readLine()) != null) {
            if (line.trim().isEmpty()) {
                continue;
            }
            final String requestLine = line;
            workers.submit(new Runnable() {
                @Override
                public void run() {
                    dispatch(requestLine);
                }
            });
        }
        System.exit(0);
    }

    private void dispatch(String line) {
        String id = "";
        try {
            Map<String, Object> request = Json.parseObject(line);
            id = Json.text(request, "id");
            String operation = Json.text(request, "op");
            if ("shutdown".equals(operation)) {
                replyOk(id);
                System.exit(0);
                return;
            }
            if ("ping".equals(operation)) {
                reply(id, "pong", message("message", "ok"));
                return;
            }
            if ("loadDriver".equals(operation)) {
                loadDriver(request, Json.text(request, "driver"));
                replyOk(id);
                return;
            }
            if ("cancel".equals(operation) || "close".equals(operation)) {
                String queryId = Json.text(request, "queryId");
                closeSession(queryId.isEmpty() ? id : queryId, true);
                replyOk(id);
                return;
            }
            if ("test".equals(operation)) {
                test(request);
                replyOk(id);
                return;
            }
            if ("catalogs".equals(operation)) {
                replyCatalogs(id, request);
                return;
            }
            if ("tables".equals(operation)) {
                replyTables(id, request);
                return;
            }
            if ("query".equals(operation)) {
                query(id, request);
                return;
            }
            replyError(id, "不支持的 JDBC 操作：" + operation);
        } catch (Exception exception) {
            replyError(id, describe(exception));
        }
    }

    private void test(Map<String, Object> request) throws Exception {
        Connection connection = openConnection(request);
        try {
            connection.getMetaData();
        } finally {
            connection.close();
        }
    }

    private void replyCatalogs(String id, Map<String, Object> request) throws Exception {
        Connection connection = openConnection(request);
        try {
            List<String> names = readCatalogNames(connection.getMetaData());
            Map<String, Object> payload = new LinkedHashMap<String, Object>();
            payload.put("names", names);
            reply(id, "catalogs", payload);
        } finally {
            connection.close();
        }
    }

    private void replyTables(String id, Map<String, Object> request) throws Exception {
        Connection connection = openConnection(request);
        try {
            String name = Json.text(request, "catalog");
            DatabaseMetaData meta = connection.getMetaData();
            boolean isCatalog = catalogExists(meta, name);
            ResultSet rows = isCatalog
                    ? meta.getTables(name, null, "%", TableTypes)
                    : meta.getTables(null, name, "%", TableTypes);
            List<Map<String, Object>> objects = new ArrayList<Map<String, Object>>();
            try {
                while (rows.next()) {
                    String tableName = rows.getString("TABLE_NAME");
                    if (tableName == null || tableName.trim().isEmpty()) {
                        continue;
                    }
                    String schema = rows.getString("TABLE_SCHEM");
                    String tableType = rows.getString("TABLE_TYPE");
                    Map<String, Object> item = new LinkedHashMap<String, Object>();
                    item.put("schema", schema == null ? "" : schema);
                    item.put("name", tableName);
                    item.put("kind", tableType != null && tableType.toUpperCase().contains("VIEW") ? "view" : "table");
                    objects.add(item);
                }
            } finally {
                rows.close();
            }
            Map<String, Object> payload = new LinkedHashMap<String, Object>();
            payload.put("objects", objects);
            reply(id, "tables", payload);
        } finally {
            connection.close();
        }
    }

    private void query(String id, Map<String, Object> request) {
        Connection connection = null;
        Statement statement = null;
        ResultSet rows = null;
        Session session = null;
        boolean failed = false;
        try {
            connection = openConnection(request);
            session = new Session(connection);
            sessions.put(id, session);
            applySelection(connection, Json.text(request, "catalog"));
            statement = connection.createStatement();
            session.statement = statement;
            int timeoutSeconds = Json.integer(request, "timeoutSeconds", 0);
            if (timeoutSeconds > 0) {
                try {
                    statement.setQueryTimeout(timeoutSeconds);
                } catch (SQLException ignored) {
                    // 部分驱动不支持查询超时，继续执行并由调用方取消。
                }
            }
            try {
                statement.setFetchSize(RowBatchSize);
            } catch (SQLException ignored) {
                // 驱动拒绝预取大小时仍按默认方式读取。
            }
            rows = statement.executeQuery(Json.text(request, "sql"));
            writeColumns(id, rows.getMetaData());
            writeRows(id, rows, session.cancelled);
        } catch (Exception exception) {
            if (session == null || !session.cancelled.get()) {
                replyError(id, describe(exception));
                failed = true;
            }
        } finally {
            sessions.remove(id);
            if (rows != null) {
                try {
                    rows.close();
                } catch (SQLException ignored) {
                }
            }
            if (statement != null) {
                try {
                    statement.close();
                } catch (SQLException ignored) {
                }
            }
            if (connection != null) {
                try {
                    connection.close();
                } catch (SQLException ignored) {
                }
            }
        }
        if (!failed) {
            reply(id, "end", null);
        }
    }

    private void writeColumns(String id, ResultSetMetaData meta) throws Exception {
        int count = meta.getColumnCount();
        List<Map<String, Object>> columns = new ArrayList<Map<String, Object>>();
        for (int index = 1; index <= count; index++) {
            String name = meta.getColumnLabel(index);
            if (name == null || name.trim().isEmpty()) {
                name = "column" + index;
            }
            Map<String, Object> column = new LinkedHashMap<String, Object>();
            column.put("name", name);
            column.put("type", clrType(meta.getColumnType(index)));
            columns.add(column);
        }
        Map<String, Object> payload = new LinkedHashMap<String, Object>();
        payload.put("columns", columns);
        reply(id, "columns", payload);
    }

    private void writeRows(String id, ResultSet rows, AtomicBoolean cancelled) throws Exception {
        ResultSetMetaData meta = rows.getMetaData();
        int count = meta.getColumnCount();
        String[] types = new String[count];
        for (int index = 0; index < count; index++) {
            types[index] = clrType(meta.getColumnType(index + 1));
        }
        List<List<Object>> batch = new ArrayList<List<Object>>();
        while (!cancelled.get() && rows.next()) {
            List<Object> row = new ArrayList<Object>();
            for (int index = 1; index <= count; index++) {
                row.add(cellValue(rows.getObject(index), types[index - 1]));
            }
            batch.add(row);
            if (batch.size() >= RowBatchSize) {
                writeRowBatch(id, batch);
                batch = new ArrayList<List<Object>>();
            }
        }
        if (!batch.isEmpty()) {
            writeRowBatch(id, batch);
        }
    }

    private void writeRowBatch(String id, List<List<Object>> rows) {
        Map<String, Object> payload = new LinkedHashMap<String, Object>();
        payload.put("rows", rows);
        reply(id, "rows", payload);
    }

    private Object cellValue(Object value, String clrType) {
        if (value == null) {
            return null;
        }
        if ("bool".equals(clrType)) {
            if (value instanceof Boolean) {
                return value;
            }
            if (value instanceof Number) {
                return ((Number) value).intValue() != 0;
            }
            return Boolean.valueOf(String.valueOf(value));
        }
        if ("int".equals(clrType) || "long".equals(clrType)) {
            if (value instanceof Number) {
                return Long.valueOf(((Number) value).longValue());
            }
            return String.valueOf(value);
        }
        if ("double".equals(clrType)) {
            if (value instanceof Number) {
                return Double.valueOf(((Number) value).doubleValue());
            }
            return String.valueOf(value);
        }
        if ("decimal".equals(clrType)) {
            if (value instanceof BigDecimal) {
                return ((BigDecimal) value).toPlainString();
            }
            return String.valueOf(value);
        }
        if (value instanceof byte[]) {
            return Base64.getEncoder().encodeToString((byte[]) value);
        }
        return String.valueOf(value);
    }

    private static String clrType(int jdbcType) {
        switch (jdbcType) {
            case Types.TINYINT:
            case Types.SMALLINT:
            case Types.INTEGER:
                return "int";
            case Types.BIGINT:
                return "long";
            case Types.FLOAT:
            case Types.REAL:
            case Types.DOUBLE:
                return "double";
            case Types.DECIMAL:
            case Types.NUMERIC:
                return "decimal";
            case Types.BIT:
            case Types.BOOLEAN:
                return "bool";
            default:
                return "string";
        }
    }

    private Connection openConnection(Map<String, Object> request) throws Exception {
        String driverName = Json.text(request, "driver");
        String url = Json.text(request, "url");
        Driver driver = loadDriver(request, driverName);
        Properties properties = new Properties();
        String user = Json.text(request, "user");
        String password = Json.text(request, "password");
        if (!user.isEmpty()) {
            properties.setProperty("user", user);
        }
        properties.setProperty("password", password);
        if (driverName.endsWith("OdpsDriver")) {
            if (!user.isEmpty()) {
                properties.setProperty("access_id", user);
            }
            properties.setProperty("access_key", password);
        }
        Connection connection = driver.connect(url, properties);
        if (connection == null) {
            throw new SQLException("驱动不接受该 JDBC URL：" + url);
        }
        try {
            connection.setReadOnly(true);
        } catch (SQLException ignored) {
            // 只读标志不是所有驱动都实现；写操作仍由插件侧的 SQL 校验拒绝。
        }
        return connection;
    }

    private Driver loadDriver(Map<String, Object> request, String driverName) throws Exception {
        if (driverName == null || driverName.trim().isEmpty()) {
            throw new SQLException("驱动类不能为空。");
        }
        List<String> jars = readJarPaths(request);
        String signature = driverName + "|" + joinJars(jars);
        Driver cached = drivers.get(signature);
        if (cached != null) {
            return cached;
        }
        ClassLoader loader = classLoaderFor(jars);
        Class<?> type;
        try {
            type = Class.forName(driverName, true, loader);
        } catch (ClassNotFoundException exception) {
            throw new SQLException("找不到驱动类 " + driverName + "。请在连接方案中添加正确的 JDBC 驱动 jar。");
        }
        Object instance = type.getDeclaredConstructor().newInstance();
        if (!(instance instanceof Driver)) {
            throw new SQLException(driverName + " 不是 JDBC 驱动类。");
        }
        Driver driver = (Driver) instance;
        drivers.put(signature, driver);
        return driver;
    }

    private ClassLoader classLoaderFor(List<String> jars) throws Exception {
        String signature = joinJars(jars);
        ClassLoader existing = loadersBySignature.get(signature);
        if (existing != null) {
            return existing;
        }
        List<java.net.URL> urls = new ArrayList<java.net.URL>();
        for (String jar : jars) {
            File file = new File(jar);
            if (!file.isFile()) {
                throw new SQLException("找不到驱动 jar：" + jar);
            }
            urls.add(file.toURI().toURL());
        }
        if (urls.isEmpty()) {
            throw new SQLException("请至少指定一个驱动 jar。");
        }
        ClassLoader loader = new java.net.URLClassLoader(urls.toArray(new java.net.URL[urls.size()]), Bridge.class.getClassLoader());
        loadersBySignature.put(signature, loader);
        return loader;
    }

    private static List<String> readJarPaths(Map<String, Object> request) {
        List<String> jars = new ArrayList<String>();
        Object value = request.get("jars");
        if (value instanceof List) {
            List<?> items = (List<?>) value;
            for (Object item : items) {
                if (item == null) continue;
                String path = String.valueOf(item).trim();
                if (!path.isEmpty()) {
                    jars.add(path);
                }
            }
        }
        return jars;
    }

    private static String joinJars(List<String> jars) {
        StringBuilder builder = new StringBuilder();
        for (String jar : jars) {
            if (builder.length() > 0) {
                builder.append('|');
            }
            builder.append(jar);
        }
        return builder.toString();
    }

    private static List<String> readCatalogNames(DatabaseMetaData meta) throws SQLException {
        List<String> names = new ArrayList<String>();
        ResultSet catalogs = meta.getCatalogs();
        try {
            while (catalogs.next()) {
                addName(names, catalogs.getString("TABLE_CAT"));
            }
        } finally {
            catalogs.close();
        }
        if (!names.isEmpty()) {
            return names;
        }
        ResultSet schemas = meta.getSchemas();
        try {
            while (schemas.next()) {
                addName(names, schemas.getString("TABLE_SCHEM"));
            }
        } finally {
            schemas.close();
        }
        return names;
    }

    private static boolean catalogExists(DatabaseMetaData meta, String name) throws SQLException {
        if (name == null || name.isEmpty()) {
            return false;
        }
        ResultSet catalogs = meta.getCatalogs();
        try {
            while (catalogs.next()) {
                if (name.equalsIgnoreCase(catalogs.getString("TABLE_CAT"))) {
                    return true;
                }
            }
        } finally {
            catalogs.close();
        }
        return false;
    }

    private static void applySelection(Connection connection, String name) {
        if (name == null || name.trim().isEmpty()) {
            return;
        }
        try {
            DatabaseMetaData meta = connection.getMetaData();
            if (catalogExists(meta, name)) {
                connection.setCatalog(name);
                return;
            }
        } catch (SQLException ignored) {
            // 继续尝试按架构切换。
        }
        try {
            connection.setSchema(name);
        } catch (SQLException ignored) {
            // 驱动不支持切换架构时，查询仍使用连接 URL 中的默认库。
        }
    }

    private static void addName(List<String> names, String name) {
        if (name == null || name.trim().isEmpty()) {
            return;
        }
        for (String existing : names) {
            if (existing.equalsIgnoreCase(name)) {
                return;
            }
        }
        names.add(name);
    }

    private void closeSession(String id, boolean cancel) {
        Session session = sessions.remove(id);
        if (session == null) {
            return;
        }
        session.cancelled.set(true);
        if (cancel && session.statement != null) {
            try {
                session.statement.cancel();
            } catch (SQLException ignored) {
            }
        }
        try {
            session.connection.close();
        } catch (SQLException ignored) {
        }
    }

    private void replyOk(String id) {
        reply(id, "ok", null);
    }

    private void replyError(String id, String message) {
        Map<String, Object> payload = message("message", message == null || message.trim().isEmpty() ? "JDBC 操作失败。" : message);
        reply(id, "error", payload);
    }

    private void reply(String id, String type, Map<String, Object> extra) {
        Map<String, Object> payload = new LinkedHashMap<String, Object>();
        payload.put("id", id == null ? "" : id);
        payload.put("type", type);
        if (extra != null) {
            payload.putAll(extra);
        }
        synchronized (outputLock) {
            try {
                output.write(Json.write(payload));
                output.write('\n');
                output.flush();
            } catch (Exception exception) {
                // 标准输出中断后调用方会结束进程，这里不能再抛回查询线程。
            }
        }
    }

    private static Map<String, Object> message(String key, String value) {
        Map<String, Object> payload = new LinkedHashMap<String, Object>();
        payload.put(key, value);
        return payload;
    }

    private static String describe(Exception exception) {
        String message = exception.getMessage();
        if (message == null || message.trim().isEmpty()) {
            return exception.getClass().getSimpleName();
        }
        return message;
    }

    private static final class Session {
        private final Connection connection;
        private final AtomicBoolean cancelled = new AtomicBoolean();
        private volatile Statement statement;

        private Session(Connection connection) {
            this.connection = connection;
        }
    }
}
