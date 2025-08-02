import java.sql.*;
import java.util.*;
import java.util.stream.Collectors;

public class DatabaseCSVMatcher {
    
    public static class MatchResult {
        private List<Map<String, Object>> matched = new ArrayList<>();
        private List<Map<String, Object>> unmatched = new ArrayList<>();
        private List<Map<String, Object>> dbOnly = new ArrayList<>();
        
        // Getters and setters
        public List<Map<String, Object>> getMatched() { return matched; }
        public List<Map<String, Object>> getUnmatched() { return unmatched; }
        public List<Map<String, Object>> getDbOnly() { return dbOnly; }
        
        public void addMatched(Map<String, Object> record) { matched.add(record); }
        public void addUnmatched(Map<String, Object> record) { unmatched.add(record); }
        public void addDbOnly(Map<String, Object> record) { dbOnly.add(record); }
    }
    
    /**
     * Compare CSV records with database records
     * @param csvRecords List of CSV records as Map
     * @param connection Database connection
     * @param tableName Target table name
     * @param primaryKeys Array of primary key column names
     * @return MatchResult containing matched, unmatched, and db-only records
     */
    public static MatchResult compareRecords(List<Map<String, Object>> csvRecords, 
                                           Connection connection, 
                                           String tableName, 
                                           String... primaryKeys) throws SQLException {
        
        MatchResult result = new MatchResult();
        
        if (csvRecords.isEmpty()) {
            return result;
        }
        
        // Step 1: Get all DB records
        Map<String, Map<String, Object>> dbRecordsMap = fetchDatabaseRecords(
            connection, tableName, primaryKeys);
        
        // Step 2: Compare CSV records with DB records
        for (Map<String, Object> csvRecord : csvRecords) {
            String compositeKey = createCompositeKey(csvRecord, primaryKeys);
            
            if (dbRecordsMap.containsKey(compositeKey)) {
                Map<String, Object> dbRecord = dbRecordsMap.get(compositeKey);
                
                // Add both CSV and DB record for comparison
                Map<String, Object> matchedRecord = new HashMap<>();
                matchedRecord.put("csv_record", csvRecord);
                matchedRecord.put("db_record", dbRecord);
                matchedRecord.put("matches", recordsMatch(csvRecord, dbRecord));
                
                result.addMatched(matchedRecord);
                dbRecordsMap.remove(compositeKey); // Remove matched record
            } else {
                result.addUnmatched(csvRecord);
            }
        }
        
        // Step 3: Remaining DB records are not in CSV
        dbRecordsMap.values().forEach(result::addDbOnly);
        
        return result;
    }
    
    /**
     * Fetch all records from database and create a map with composite keys
     */
    private static Map<String, Map<String, Object>> fetchDatabaseRecords(
            Connection connection, String tableName, String[] primaryKeys) throws SQLException {
        
        Map<String, Map<String, Object>> dbRecords = new HashMap<>();
        
        String sql = "SELECT * FROM " + tableName;
        
        try (PreparedStatement stmt = connection.prepareStatement(sql);
             ResultSet rs = stmt.executeQuery()) {
            
            ResultSetMetaData metaData = rs.getMetaData();
            int columnCount = metaData.getColumnCount();
            
            while (rs.next()) {
                Map<String, Object> record = new HashMap<>();
                
                // Get all columns
                for (int i = 1; i <= columnCount; i++) {
                    String columnName = metaData.getColumnName(i);
                    Object value = rs.getObject(i);
                    record.put(columnName.toLowerCase(), value);
                }
                
                String compositeKey = createCompositeKey(record, primaryKeys);
                dbRecords.put(compositeKey, record);
            }
        }
        
        return dbRecords;
    }
    
    /**
     * Create composite key from primary key values
     */
    private static String createCompositeKey(Map<String, Object> record, String[] primaryKeys) {
        return Arrays.stream(primaryKeys)
                .map(key -> String.valueOf(record.get(key.toLowerCase())))
                .collect(Collectors.joining("||"));
    }
    
    /**
     * Check if two records match (excluding primary keys)
     */
    private static boolean recordsMatch(Map<String, Object> csvRecord, Map<String, Object> dbRecord) {
        for (Map.Entry<String, Object> entry : csvRecord.entrySet()) {
            String key = entry.getKey().toLowerCase();
            Object csvValue = entry.getValue();
            Object dbValue = dbRecord.get(key);
            
            if (!Objects.equals(csvValue, dbValue)) {
                return false;
            }
        }
        return true;
    }
}