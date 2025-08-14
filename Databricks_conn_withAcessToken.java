import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.sql.Connection;
import java.sql.DriverManager;
import java.util.HashMap;
import java.util.Map;

public class DatabricksSPNConnect {

    // === SPN Credentials ===
    private static final String TENANT_ID = "<tenant-id>";
    private static final String CLIENT_ID = "<service-principal-app-id>";
    private static final String CLIENT_SECRET = "<service-principal-secret>";

    // === Databricks JDBC Connection Params ===
    private static final String WORKSPACE_ID = "<workspace-id>"; // e.g. adb-1234567890123456.7.azuredatabricks.net
    private static final String REGION = "<region>"; // e.g. eastus
    private static final String ORG_ID = "<org-id>"; // from workspace URL: o/<org-id>/
    private static final String CLUSTER_ID = "<cluster-id>"; // from cluster URL

    public static void main(String[] args) throws Exception {

        // Step 1: Fetch OAuth2 token for Databricks
        String accessToken = getDatabricksToken();
        System.out.println("Access Token fetched successfully.");

        // Step 2: Build JDBC connection string
        String jdbcUrl =
                "jdbc:databricks://adb-" + WORKSPACE_ID + "." + REGION + ".azuredatabricks.net:443/default;" +
                "ssl=1;" +
                "transportMode=http;" +
                "AuthMech=3;" +
                "UID=token;" +
                "PWD=" + accessToken + ";" +
                "HTTPPath=/sql/protocolv1/o/" + ORG_ID + "/" + CLUSTER_ID + ";";

        // Step 3: Connect using Databricks JDBC driver
        Class.forName("com.databricks.client.jdbc.Driver");

        try (Connection conn = DriverManager.getConnection(jdbcUrl)) {
            System.out.println("✅ Connected to Databricks cluster successfully.");
        }
    }

    private static String getDatabricksToken() throws Exception {
        String tokenEndpoint = "https://login.microsoftonline.com/" + TENANT_ID + "/oauth2/token";

        Map<Object, Object> data = new HashMap<>();
        data.put("grant_type", "client_credentials");
        data.put("client_id", CLIENT_ID);
        data.put("client_secret", CLIENT_SECRET);
        data.put("resource", "https://databricks.azure.net");

        HttpClient client = HttpClient.newHttpClient();
        HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(tokenEndpoint))
                .header("Content-Type", "application/x-www-form-urlencoded")
                .POST(ofFormData(data))
                .build();

        HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

        if (response.statusCode() != 200) {
            throw new RuntimeException("Failed to get token: " + response.body());
        }

        String body = response.body();
        String accessToken = body.split("\"access_token\"\\s*:\\s*\"")[1].split("\"")[0];
        return accessToken;
    }

    private static HttpRequest.BodyPublisher ofFormData(Map<Object, Object> data) {
        StringBuilder builder = new StringBuilder();
        for (Map.Entry<Object, Object> entry : data.entrySet()) {
            if (builder.length() > 0) {
                builder.append("&");
            }
            builder.append(entry.getKey()).append("=")
                    .append(java.net.URLEncoder.encode(entry.getValue().toString(), java.nio.charset.StandardCharsets.UTF_8));
        }
        return HttpRequest.BodyPublishers.ofString(builder.toString());
    }
}
