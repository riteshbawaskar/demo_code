import java.sql.*;
import java.util.Properties;

public class DatabricksConnection {
    
    // Databricks connection parameters
    private static final String DATABRICKS_HOST = "your-workspace-url"; // e.g., "adb-123456789.12.azuredatabricks.net"
    private static final String HTTP_PATH = "/sql/1.0/warehouses/your-warehouse-id"; // e.g., "/sql/1.0/warehouses/abc123def456"
    private static final String CLIENT_ID = "your-service-principal-client-id";
    private static final String CLIENT_SECRET = "your-service-principal-secret";
    private static final String TENANT_ID = "your-azure-tenant-id";
    
    public static void main(String[] args) {
        Connection connection = null;
        
        try {
            // Load the Databricks JDBC driver
            Class.forName("com.databricks.client.jdbc.Driver");
            
            // Validate HTTP path format
            if (!HTTP_PATH.startsWith("/")) {
                throw new IllegalArgumentException("HTTP_PATH must start with '/'");
            }
            
            // Method 1: M2M (Machine-to-Machine) OAuth flow - Recommended for SPN
            Properties props = new Properties();
            props.setProperty("AuthMech", "11"); // OAuth 2.0
            props.setProperty("Auth_Flow", "2"); // Changed to 2 for M2M flow instead of 1
            props.setProperty("OAuth2ClientId", CLIENT_ID);
            props.setProperty("OAuth2Secret", CLIENT_SECRET);
            props.setProperty("OAuth2TenantId", TENANT_ID);
            props.setProperty("HTTPPath", HTTP_PATH);
            props.setProperty("SSL", "1");
            // Add Azure-specific OAuth settings
            props.setProperty("OAuth2TokenEndpoint", String.format("https://login.microsoftonline.com/%s/oauth2/v2.0/token", TENANT_ID));
            props.setProperty("OAuth2Scope", "2ff814a6-3304-4ab8-85cb-cd0e6f879c1d/.default"); // Databricks scope
            
            String baseUrl = String.format("jdbc:databricks://%s:443", DATABRICKS_HOST);
            
            System.out.println("Attempting M2M OAuth connection...");
            System.out.println("Base URL: " + baseUrl);
            System.out.println("HTTP Path: " + HTTP_PATH);
            System.out.println("Tenant ID: " + TENANT_ID);
            
            connection = DriverManager.getConnection(baseUrl, props);
            System.out.println("Connected to Databricks successfully!");
            
            // Alternative Method 2: Personal Access Token approach (uncomment if OAuth fails)
            /*
            // If you have a PAT instead of SPN, use this:
            Properties patProps = new Properties();
            patProps.setProperty("AuthMech", "3"); // Personal Access Token
            patProps.setProperty("PWD", "your-personal-access-token-here");
            patProps.setProperty("HTTPPath", HTTP_PATH);
            patProps.setProperty("SSL", "1");
            
            connection = DriverManager.getConnection(baseUrl, patProps);
            */
            
            // Alternative Method 3: Azure AD integrated (uncomment for testing)
            /*
            Properties aadProps = new Properties();
            aadProps.setProperty("AuthMech", "11");
            aadProps.setProperty("Auth_Flow", "1"); // Client credentials
            aadProps.setProperty("OAuth2ClientId", CLIENT_ID);
            aadProps.setProperty("OAuth2Secret", CLIENT_SECRET);
            // Explicit Azure AD authority
            aadProps.setProperty("OAuth2AuthorityUrl", String.format("https://login.microsoftonline.com/%s", TENANT_ID));
            aadProps.setProperty("HTTPPath", HTTP_PATH);
            aadProps.setProperty("SSL", "1");
            
            connection = DriverManager.getConnection(baseUrl, aadProps);
            */
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
        <version>2.6.36</version>
    </dependency>
</dependencies>

Gradle Dependencies (add to build.gradle):
dependencies {
    implementation 'com.databricks:databricks-jdbc:2.6.36'
}

Configuration Steps:
1. Replace placeholder values:
   - DATABRICKS_HOST: Your workspace URL (without https://)
     Example: "adb-1234567890123456.12.azuredatabricks.net"
   - HTTP_PATH: Path to your SQL warehouse or cluster
   - CLIENT_ID: Service Principal Application ID
   - CLIENT_SECRET: Service Principal secret  
   - TENANT_ID: Azure AD tenant ID (GUID format)

2. Service Principal Setup Requirements:
   - SPN must be registered in Azure AD
   - SPN must be added to Databricks workspace as a user
   - SPN needs "Contributor" role or higher on Databricks workspace
   - SPN needs "Can Use" permission on SQL Warehouse/Cluster

3. Common Authorization URL Error Fixes:
   - Verify TENANT_ID is correct (should be GUID, not domain name)
   - Ensure CLIENT_ID and CLIENT_SECRET are valid and not expired
   - Check that SPN has proper API permissions in Azure AD:
     * Azure Databricks (2ff814a6-3304-4ab8-85cb-cd0e6f879c1d)
     * Grant admin consent for the permissions
   - Try Auth_Flow=2 (M2M) instead of Auth_Flow=1 (client credentials)

4. Alternative if OAuth continues failing:
   - Generate a Personal Access Token in Databricks
   - Use AuthMech=3 with the PAT (see Alternative Method 2 in code)

5. Databricks Workspace Configuration:
   - Go to Admin Console → Identity and Access → Service Principals
   - Add your SPN and assign appropriate permissions
   - Ensure the workspace allows service principal authentication

6. Network/Firewall:
   - Ensure outbound access to login.microsoftonline.com
   - Verify Databricks workspace URL is accessible
   - Check if corporate firewall blocks OAuth endpoints
*/