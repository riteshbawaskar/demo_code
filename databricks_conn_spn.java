import java.sql.*;
import java.util.Properties;

public class DatabricksConnection {
    
    // Databricks connection parameters
    private static final String DATABRICKS_HOST = "your-workspace-url"; // e.g., "adb-123456789.azuredatabricks.net"
    private static final String HTTP_PATH = "/sql/1.0/warehouses/your-warehouse-id"; // or cluster path
    private static final String CLIENT_ID = "your-service-principal-client-id";
    private static final String CLIENT_SECRET = "your-service-principal-secret";
    private static final String TENANT_ID = "your-azure-tenant-id";
    
    public static void main(String[] args) {
        Connection connection = null;
        
        try {
            // Load the Databricks JDBC driver
            Class.forName("com.databricks.client.jdbc.Driver");
            
            // Build connection URL
            String jdbcUrl = String.format(
                "jdbc:databricks://%s:443%s;AuthMech=11;Auth_Flow=1;OAuth2ClientId=%s;OAuth2Secret=%s;OAuth2TenantId=%s",
                DATABRICKS_HOST,
                HTTP_PATH,
                CLIENT_ID,
                CLIENT_SECRET,
                TENANT_ID
            );
            
            // Alternative: using Properties for cleaner configuration
            Properties props = new Properties();
            props.setProperty("AuthMech", "11"); // OAuth 2.0
            props.setProperty("Auth_Flow", "1"); // Client credentials flow
            props.setProperty("OAuth2ClientId", CLIENT_ID);
            props.setProperty("OAuth2Secret", CLIENT_SECRET);
            props.setProperty("OAuth2TenantId", TENANT_ID);
            
            String baseUrl = String.format("jdbc:databricks://%s:443%s", DATABRICKS_HOST, HTTP_PATH);
            
            // Establish connection
            connection = DriverManager.getConnection(baseUrl, props);
            System.out.println("Connected to Databricks successfully!");
            
            // Query tables
            queryTables(connection);
            
        } catch (ClassNotFoundException e) {
            System.err.println("Databricks JDBC driver not found: " + e.getMessage());
        } catch (SQLException e) {
            System.err.println("Database connection error: " + e.getMessage());
        } finally {
            // Close connection
            if (connection != null) {
                try {
                    connection.close();
                    System.out.println("Connection closed.");
                } catch (SQLException e) {
                    System.err.println("Error closing connection: " + e.getMessage());
                }
            }
        }
    }
    
    private static void queryTables(Connection connection) throws SQLException {
        // Example 1: List all tables in default database
        System.out.println("\n=== Available Tables ===");
        String showTablesQuery = "SHOW TABLES";
        executeQuery(connection, showTablesQuery);
        
        // Example 2: Query specific table
        System.out.println("\n=== Sample Data Query ===");
        String dataQuery = "SELECT * FROM your_table_name LIMIT 10";
        executeQuery(connection, dataQuery);
        
        // Example 3: Count records in a table
        System.out.println("\n=== Record Count ===");
        String countQuery = "SELECT COUNT(*) as total_records FROM your_table_name";
        executeQuery(connection, countQuery);
    }
    
    private static void executeQuery(Connection connection, String sql) throws SQLException {
        try (Statement statement = connection.createStatement();
             ResultSet resultSet = statement.executeQuery(sql)) {
            
            // Get column metadata
            ResultSetMetaData metaData = resultSet.getMetaData();
            int columnCount = metaData.getColumnCount();
            
            // Print column headers
            for (int i = 1; i <= columnCount; i++) {
                System.out.print(metaData.getColumnName(i) + "\t");
            }
            System.out.println();
            
            // Print data rows
            while (resultSet.next()) {
                for (int i = 1; i <= columnCount; i++) {
                    System.out.print(resultSet.getString(i) + "\t");
                }
                System.out.println();
            }
            System.out.println("Query executed successfully.\n");
        }
    }
}

/*
Maven Dependencies (add to pom.xml):
<dependencies>
    <dependency>
        <groupId>com.databricks</groupId>
        <artifactId>databricks-jdbc</artifactId>
        <version>2.6.34</version>
    </dependency>
</dependencies>

Gradle Dependencies (add to build.gradle):
dependencies {
    implementation 'com.databricks:databricks-jdbc:2.6.34'
}

Configuration Steps:
1. Replace placeholder values:
   - DATABRICKS_HOST: Your workspace URL (without https://)
   - HTTP_PATH: Path to your SQL warehouse or cluster
   - CLIENT_ID: Service Principal Application ID
   - CLIENT_SECRET: Service Principal secret
   - TENANT_ID: Azure AD tenant ID

2. Ensure your Service Principal has proper permissions:
   - Contributor or higher role on the Databricks workspace
   - Access to the specific databases/tables you want to query

3. For SQL Warehouse, use path like: /sql/1.0/warehouses/warehouse-id
   For cluster, use path like: /sql/protocolv1/o/workspace-id/cluster-id
*/