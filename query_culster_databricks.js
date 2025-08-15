// cluster-query.js - Query tables from Databricks Cluster (not SQL Warehouse)

const axios = require('axios');

// Configuration
const config = {
    // Azure AD / Service Principal
    tenantId: 'your-tenant-id',
    clientId: 'your-service-principal-client-id',
    clientSecret: 'your-service-principal-secret',
    
    // Databricks
    databricksHost: 'https://your-workspace.cloud.databricks.com',
    
    // Cluster ID from HTTP Path: /sql/protocolv1/o/{workspace-id}/{cluster-id}
    clusterId: '0123-456789-abcd123'  // <-- Extract this from your HTTP Path
};

/**
 * Method 1: Using Command Execution API 1.2 (for running clusters)
 * This executes commands directly on the cluster
 */
async function queryTablesUsingCommand(token) {
    console.log('\n📊 Method 1: Using Command Execution API');
    console.log('=' .repeat(50));
    
    try {
        // Create execution context
        const contextResponse = await axios.post(
            `${config.databricksHost}/api/1.2/contexts/create`,
            {
                clusterId: config.clusterId,
                language: 'sql'
            },
            {
                headers: { 'Authorization': `Bearer ${token}` }
            }
        );
        
        const contextId = contextResponse.data.id;
        console.log(`✅ Context created: ${contextId}`);
        
        // Execute SQL command
        const commandResponse = await axios.post(
            `${config.databricksHost}/api/1.2/commands/execute`,
            {
                clusterId: config.clusterId,
                contextId: contextId,
                language: 'sql',
                command: 'SHOW TABLES'
            },
            {
                headers: { 'Authorization': `Bearer ${token}` }
            }
        );
        
        const commandId = commandResponse.data.id;
        console.log(`✅ Command submitted: ${commandId}`);
        
        // Get command status and results
        let result = null;
        for (let i = 0; i < 30; i++) {
            const statusResponse = await axios.get(
                `${config.databricksHost}/api/1.2/commands/status`,
                {
                    params: {
                        clusterId: config.clusterId,
                        contextId: contextId,
                        commandId: commandId
                    },
                    headers: { 'Authorization': `Bearer ${token}` }
                }
            );
            
            if (statusResponse.data.status === 'Finished') {
                result = statusResponse.data.results;
                break;
            } else if (statusResponse.data.status === 'Error') {
                console.error('❌ Command failed:', statusResponse.data.results);
                return null;
            }
            
            process.stdout.write('.');
            await new Promise(r => setTimeout(r, 1000));
        }
        
        // Clean up context
        await axios.post(
            `${config.databricksHost}/api/1.2/contexts/destroy`,
            {
                clusterId: config.clusterId,
                contextId: contextId
            },
            {
                headers: { 'Authorization': `Bearer ${token}` }
            }
        );
        
        return result;
        
    } catch (error) {
        console.error('❌ Command execution failed:', error.response?.data || error.message);
        
        if (error.response?.status === 404) {
            console.log('\n💡 Cluster might be stopped or doesn\'t exist');
            console.log('   Start the cluster in Databricks UI first');
        }
        
        return null;
    }
}

/**
 * Method 2: Using Jobs API to run SQL on cluster
 * This creates a one-time job to execute SQL
 */
async function queryTablesUsingJob(token) {
    console.log('\n📊 Method 2: Using Jobs API');
    console.log('=' .repeat(50));
    
    try {
        // Create a one-time job
        const jobResponse = await axios.post(
            `${config.databricksHost}/api/2.1/jobs/runs/submit`,
            {
                run_name: 'Query Tables via API',
                existing_cluster_id: config.clusterId,
                tasks: [
                    {
                        task_key: 'query_tables',
                        sql_task: {
                            query: {
                                query: 'SHOW TABLES'
                            }
                        }
                    }
                ]
            },
            {
                headers: { 'Authorization': `Bearer ${token}` }
            }
        );
        
        const runId = jobResponse.data.run_id;
        console.log(`✅ Job submitted: Run ID ${runId}`);
        
        // Wait for job completion
        let jobResult = null;
        for (let i = 0; i < 60; i++) {
            const statusResponse = await axios.get(
                `${config.databricksHost}/api/2.1/jobs/runs/get`,
                {
                    params: { run_id: runId },
                    headers: { 'Authorization': `Bearer ${token}` }
                }
            );
            
            const state = statusResponse.data.state.life_cycle_state;
            
            if (state === 'TERMINATED') {
                const resultState = statusResponse.data.state.result_state;
                if (resultState === 'SUCCESS') {
                    console.log('✅ Job completed successfully');
                    
                    // Get the output
                    const outputResponse = await axios.get(
                        `${config.databricksHost}/api/2.1/jobs/runs/get-output`,
                        {
                            params: { run_id: runId },
                            headers: { 'Authorization': `Bearer ${token}` }
                        }
                    );
                    
                    jobResult = outputResponse.data;
                } else {
                    console.error('❌ Job failed:', statusResponse.data.state.state_message);
                }
                break;
            }
            
            process.stdout.write('.');
            await new Promise(r => setTimeout(r, 1000));
        }
        
        return jobResult;
        
    } catch (error) {
        console.error('❌ Job execution failed:', error.response?.data || error.message);
        return null;
    }
}

/**
 * Method 3: Using Databricks SQL Connector approach (simpler for JDBC-like access)
 * This uses a different endpoint that works with clusters
 */
async function queryTablesUsingSQLEndpoint(token) {
    console.log('\n📊 Method 3: Using Statement API with Cluster');
    console.log('=' .repeat(50));
    
    try {
        // Some Databricks deployments support this
        const response = await axios.post(
            `${config.databricksHost}/api/2.0/sql/statements`,
            {
                cluster_id: config.clusterId,  // Note: using cluster_id instead of warehouse_id
                statement: 'SHOW TABLES',
                wait_timeout: '10s',
                on_wait_timeout: 'CONTINUE'
            },
            {
                headers: { 'Authorization': `Bearer ${token}` }
            }
        );
        
        const statementId = response.data.statement_id;
        console.log(`✅ Statement created: ${statementId}`);
        
        // Get results
        let result = null;
        for (let i = 0; i < 30; i++) {
            const statusResponse = await axios.get(
                `${config.databricksHost}/api/2.0/sql/statements/${statementId}`,
                { headers: { 'Authorization': `Bearer ${token}` } }
            );
            
            if (statusResponse.data.status.state === 'SUCCEEDED') {
                result = statusResponse.data.result;
                break;
            }
            
            await new Promise(r => setTimeout(r, 1000));
        }
        
        return result;
        
    } catch (error) {
        if (error.response?.status === 400) {
            console.log('❌ This Databricks deployment doesn\'t support cluster_id with SQL statements API');
            console.log('   Use Method 1 (Command API) or Method 2 (Jobs API) instead');
        } else {
            console.error('❌ Error:', error.response?.data?.message || error.message);
        }
        return null;
    }
}

/**
 * Helper function to check cluster status
 */
async function checkClusterStatus(token) {
    try {
        const response = await axios.get(
            `${config.databricksHost}/api/2.0/clusters/get`,
            {
                params: { cluster_id: config.clusterId },
                headers: { 'Authorization': `Bearer ${token}` }
            }
        );
        
        const cluster = response.data;
        console.log('\n📊 Cluster Information:');
        console.log('=' .repeat(50));
        console.log(`Name: ${cluster.cluster_name}`);
        console.log(`ID: ${cluster.cluster_id}`);
        console.log(`State: ${cluster.state}`);
        console.log(`Spark Version: ${cluster.spark_version}`);
        console.log(`Driver Node: ${cluster.driver_node_type_id}`);
        
        if (cluster.state !== 'RUNNING') {
            console.log('\n⚠️  Cluster is not running!');
            console.log('   Start the cluster first or use a SQL Warehouse instead');
            return false;
        }
        
        return true;
        
    } catch (error) {
        console.error('❌ Could not get cluster info:', error.response?.data?.message || error.message);
        return false;
    }
}

/**
 * Main function - Try all methods
 */
async function queryClusterTables() {
    try {
        // Step 1: Get OAuth token
        console.log('🔐 Getting OAuth token...');
        const tokenResponse = await axios.post(
            `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
            new URLSearchParams({
                'grant_type': 'client_credentials',
                'client_id': config.clientId,
                'client_secret': config.clientSecret,
                'scope': '2ff814a6-3304-4ab8-85cb-cd0e6f879c1d/.default'
            })
        );
        const token = tokenResponse.data.access_token;
        console.log('✅ Token obtained');
        
        // Step 2: Check cluster status
        const clusterRunning = await checkClusterStatus(token);
        if (!clusterRunning) {
            console.log('\n💡 Solutions:');
            console.log('1. Start the cluster in Databricks UI');
            console.log('2. Use a SQL Warehouse instead (recommended for SQL queries)');
            console.log('3. Create a job cluster that starts automatically');
            return;
        }
        
        // Step 3: Try different methods
        console.log('\n🔍 Attempting to query tables...\n');
        
        // Try Method 1: Command API (most reliable for clusters)
        let result = await queryTablesUsingCommand(token);
        
        if (result && result.data) {
            console.log('\n✅ Tables found using Command API:');
            console.log(result.data);
            return;
        }
        
        // Try Method 2: Jobs API
        result = await queryTablesUsingJob(token);
        
        if (result) {
            console.log('\n✅ Tables found using Jobs API:');
            console.log(JSON.stringify(result, null, 2));
            return;
        }
        
        // Try Method 3: SQL Statement API with cluster_id
        result = await queryTablesUsingSQLEndpoint(token);
        
        if (result) {
            console.log('\n✅ Tables found using SQL Statement API:');
            result.data_array.forEach((row, i) => {
                console.log(`${i + 1}. ${row[0]}.${row[1]}`);
            });
            return;
        }
        
        console.log('\n❌ Could not query tables from cluster');
        console.log('\n💡 Recommendations:');
        console.log('1. Use a SQL Warehouse instead of a cluster for SQL queries');
        console.log('2. Ensure the cluster has the necessary libraries installed');
        console.log('3. Check that your Service Principal has cluster attach permissions');
        
    } catch (error) {
        console.error('❌ Error:', error.message);
    }
}

/**
 * Alternative: List all clusters to find the right one
 */
async function listClusters() {
    try {
        const tokenResponse = await axios.post(
            `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
            new URLSearchParams({
                'grant_type': 'client_credentials',
                'client_id': config.clientId,
                'client_secret': config.clientSecret,
                'scope': '2ff814a6-3304-4ab8-85cb-cd0e6f879c1d/.default'
            })
        );
        const token = tokenResponse.data.access_token;
        
        const response = await axios.get(
            `${config.databricksHost}/api/2.0/clusters/list`,
            { headers: { 'Authorization': `Bearer ${token}` } }
        );
        
        console.log('\n📋 Available Clusters:');
        console.log('=' .repeat(50));
        
        const clusters = response.data.clusters || [];
        clusters.forEach(cluster => {
            console.log(`\nName: ${cluster.cluster_name}`);
            console.log(`ID: ${cluster.cluster_id}`);
            console.log(`State: ${cluster.state}`);
            console.log(`Type: ${cluster.cluster_source || 'UNKNOWN'}`);
        });
        
        if (clusters.length === 0) {
            console.log('No clusters found');
        }
        
    } catch (error) {
        console.error('Error listing clusters:', error.message);
    }
}

// Run the main function
queryClusterTables();

// Uncomment to list all clusters:
// listClusters();

module.exports = {
    queryTablesUsingCommand,
    queryTablesUsingJob,
    checkClusterStatus,
    listClusters
};