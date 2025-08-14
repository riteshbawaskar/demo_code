import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

public class DatabricksSPNConnect {
    public static void main(String[] args) throws Exception {
        Class.forName("com.databricks.client.jdbc.Driver");

        String hostname = "<workspace>.azuredatabricks.net";
        String httpPath = "<cluster-sql-warehouse-http-path>"; // from cluster settings
        String tenantId = "<tenant-id>";
        String clientId = "<service-principal-client-id>";
        String clientSecret = "<service-principal-client-secret>";

        String jdbcUrl = String.format(
            "jdbc:databricks://%s:443/default;" +
            "transportMode=http;ssl=1;" +
            "AuthMech=13;" +
            "UID=%s;" +
            "PWD=%s;" +
            "UseNativeQuery=1;" +
            "HTTPPath=%s;" +
            "OAuthClientID=%s;" +
            "OAuthClientSecret=%s;" +
            "OAuth2Client=azure-service-principal;" +
            "OAuth2Endpoint=https://login.microsoftonline.com/%s/oauth2/v2.0/token",
            hostname, clientId, clientSecret, httpPath, clientId, clientSecret, tenantId
        );

        try (Connection conn = DriverManager.getConnection(jdbcUrl);
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery("SELECT current_date()")) {

            while (rs.next()) {
                System.out.println("Current Date: " + rs.getString(1));
            }
        }
    }
}