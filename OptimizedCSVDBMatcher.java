public class OptimizedDatabaseMatcher {
    
    /**
     * Compare using batch IN queries (more efficient for large datasets)
     */
    public static MatchResult compareRecordsOptimized(List<Map<String, Object>> csvRecords, 
                                                    Connection connection, 
                                                    String tableName, 
                                                    String... primaryKeys) throws SQLException {
        
        MatchResult result = new MatchResult();
        
        if (csvRecords.isEmpty()) {
            return result;
        }
        
        // Group CSV records by composite key
        Map<String, Map<String, Object>> csvRecordsMap = csvRecords.stream()
                .collect(Collectors.toMap(
                    record -> createCompositeKey(record, primaryKeys),
                    record -> record,
                    (existing, replacement) -> existing // Handle duplicates
                ));
        
        // Fetch matching DB records using batch query
        Map<String, Map<String, Object>> dbRecordsMap = fetchRecordsByKeys(
            connection, tableName, csvRecordsMap.keySet(), primaryKeys);
        
        // Compare records
        for (Map.Entry<String, Map<String, Object>> csvEntry : csvRecordsMap.entrySet()) {
            String compositeKey = csvEntry.getKey();
            Map<String, Object> csvRecord = csvEntry.getValue();
            
            if (dbRecordsMap.containsKey(compositeKey)) {
                Map<String, Object> dbRecord = dbRecordsMap.get(compositeKey);
                
                Map<String, Object> matchedRecord = new HashMap<>();
                matchedRecord.put("csv_record", csvRecord);
                matchedRecord.put("db_record", dbRecord);
                matchedRecord.put("matches", recordsMatch(csvRecord, dbRecord));
                matchedRecord.put("differences", findDifferences(csvRecord, dbRecord));
                
                result.addMatched(matchedRecord);
            } else {
                result.addUnmatched(csvRecord);
            }
        }
        
        return result;
    }
    
    /**
     * Fetch specific records using IN clause with composite keys
     */
    private static Map<String, Map<String, Object>> fetchRecordsByKeys(
            Connection connection, String tableName, Set<String> compositeKeys, String[] primaryKeys) throws SQLException {
        
        Map<String, Map<String, Object>> dbRecords = new HashMap<>();
        
        if (compositeKeys.isEmpty()) {
            return dbRecords;
        }
        
        // Build WHERE clause for composite keys
        String whereClause = buildWhereClause(compositeKeys, primaryKeys);
        String sql = "SELECT * FROM " + tableName + " WHERE " + whereClause;
        
        try (PreparedStatement stmt = connection.prepareStatement(sql)) {
            
            // Set parameters for composite keys
            int paramIndex = 1;
            for (String compositeKey : compositeKeys) {
                String[] keyParts = compositeKey.split("\\|\\|");
                for (String keyPart : keyParts) {
                    stmt.setObject(paramIndex++, keyPart);
                }
            }
            
            try (ResultSet rs = stmt.executeQuery()) {
                ResultSetMetaData metaData = rs.getMetaData();
                int columnCount = metaData.getColumnCount();
                
                while (rs.next()) {
                    Map<String, Object> record = new HashMap<>();
                    
                    for (int i = 1; i <= columnCount; i++) {
                        String columnName = metaData.getColumnName(i);
                        Object value = rs.getObject(i);
                        record.put(columnName.toLowerCase(), value);
                    }
                    
                    String compositeKey = createCompositeKey(record, primaryKeys);
                    dbRecords.put(compositeKey, record);
                }
            }
        }
        
        return dbRecords;
    }
    
    /**
     * Build WHERE clause for multiple composite keys
     */
    private static String buildWhereClause(Set<String> compositeKeys, String[] primaryKeys) {
        if (primaryKeys.length == 1) {
            // Single primary key - use simple IN clause
            String placeholders = compositeKeys.stream()
                    .map(key -> "?")
                    .collect(Collectors.joining(","));
            return primaryKeys[0] + " IN (" + placeholders + ")";
        } else {
            // Multiple primary keys - use OR conditions
            String condition = Arrays.stream(primaryKeys)
                    .map(key -> key + " = ?")
                    .collect(Collectors.joining(" AND "));
            
            String orConditions = compositeKeys.stream()
                    .map(key -> "(" + condition + ")")
                    .collect(Collectors.joining(" OR "));
            
            return orConditions;
        }
    }
    
    /**
     * Find differences between CSV and DB records
     */
    private static Map<String, Object> findDifferences(Map<String, Object> csvRecord, 
                                                      Map<String, Object> dbRecord) {
        Map<String, Object> differences = new HashMap<>();
        
        for (Map.Entry<String, Object> entry : csvRecord.entrySet()) {
            String key = entry.getKey().toLowerCase();
            Object csvValue = entry.getValue();
            Object dbValue = dbRecord.get(key);
            
            if (!Objects.equals(csvValue, dbValue)) {
                Map<String, Object> diff = new HashMap<>();
                diff.put("csv_value", csvValue);
                diff.put("db_value", dbValue);
                differences.put(key, diff);
            }
        }
        
        return differences;
    }
}