// minimal-query.js - Absolute minimum code to get tables using SPN

const axios = require('axios');

// Your configuration
const tenantId = 'your-tenant-id';
const clientId = 'your-client-id';
const clientSecret = 'your-client-secret';
const databricksHost = 'https://your-workspace.cloud.databricks.com';
const warehouseId = 'your-warehouse-id'; // From HTTP Path: /sql/1.0/warehouses/{this-id}

async function queryTables() {
    // 1. Get token
    const tokenResponse = await axios.post(
        `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
        new URLSearchParams({
            'grant_type': 'client_credentials',
            'client_id': clientId,
            'client_secret': clientSecret,
            'scope': '2ff814a6-3304-4ab8-85cb-cd0e6f879c1d/.default'
        })
    );
    const token = tokenResponse.data.access_token;
    console.log('✅ Got token');

    // 2. Execute query
    const queryResponse = await axios.post(
        `${databricksHost}/api/2.0/sql/statements`,
        {
            warehouse_id: warehouseId,
            statement: 'SHOW TABLES',
            wait_timeout: '10s'
        },
        {
            headers: { 'Authorization': `Bearer ${token}` }
        }
    );
    const statementId = queryResponse.data.statement_id;
    console.log('✅ Query submitted');

    // 3. Get results (with retry logic)
    let result;
    for (let i = 0; i < 30; i++) {
        const statusResponse = await axios.get(
            `${databricksHost}/api/2.0/sql/statements/${statementId}`,
            { headers: { 'Authorization': `Bearer ${token}` } }
        );
        
        if (statusResponse.data.status.state === 'SUCCEEDED') {
            result = statusResponse.data.result;
            break;
        }
        await new Promise(r => setTimeout(r, 1000)); // Wait 1 second
    }

    // 4. Display tables
    console.log('\n📋 Tables:');
    result.data_array.forEach(row => {
        console.log(`  ${row[0]}.${row[1]}`); // database.table
    });
}

// Run it
queryTables().catch(console.error);