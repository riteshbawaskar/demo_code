// test-authentication.js - Test Service Principal OAuth authentication
require('dotenv').config();
const axios = require('axios');
const chalk = require('chalk');

const config = {
    // Azure AD / OAuth settings
    tenantId: process.env.AZURE_TENANT_ID,
    clientId: process.env.SERVICE_PRINCIPAL_CLIENT_ID,
    clientSecret: process.env.SERVICE_PRINCIPAL_SECRET,
    
    // Databricks settings
    databricksHost: process.env.DATABRICKS_HOST,
    scope: '2ff814a6-3304-4ab8-85cb-cd0e6f879c1d/.default', // Default Azure Databricks scope
};

/**
 * Test OAuth authentication flow
 */
async function testAuthentication() {
    console.log(chalk.blue.bold('\n🔐 Testing Service Principal Authentication\n'));
    
    // Step 1: Get OAuth token from Azure AD
    console.log(chalk.yellow('1. Requesting OAuth token from Azure AD...'));
    
    const tokenEndpoint = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
    
    try {
        const tokenResponse = await axios.post(
            tokenEndpoint,
            new URLSearchParams({
                'grant_type': 'client_credentials',
                'client_id': config.clientId,
                'client_secret': config.clientSecret,
                'scope': config.scope
            }),
            {
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            }
        );
        
        if (tokenResponse.data.access_token) {
            console.log(chalk.green('✅ OAuth token obtained successfully!'));
            console.log(chalk.gray(`   Token type: ${tokenResponse.data.token_type}`));
            console.log(chalk.gray(`   Expires in: ${tokenResponse.data.expires_in} seconds`));
            
            const token = tokenResponse.data.access_token;
            
            // Step 2: Test token with Databricks API
            console.log(chalk.yellow('\n2. Testing token with Databricks API...'));
            
            try {
                const dbResponse = await axios.get(
                    `${config.databricksHost}/api/2.0/clusters/list`,
                    {
                        headers: {
                            'Authorization': `Bearer ${token}`
                        },
                        validateStatus: () => true // Don't throw on any status
                    }
                );
                
                if (dbResponse.status === 200) {
                    console.log(chalk.green('✅ Successfully authenticated with Databricks!'));
                    console.log(chalk.gray('   The Service Principal can access Databricks APIs'));
                } else if (dbResponse.status === 403) {
                    console.log(chalk.yellow('⚠️  Token is valid but lacks Databricks permissions'));
                    console.log(chalk.gray('   The SPN may need additional permissions in Databricks'));
                } else if (dbResponse.status === 401) {
                    console.log(chalk.red('❌ Authentication failed with Databricks'));
                    console.log(chalk.gray('   The token may not be valid for this workspace'));
                } else {
                    console.log(chalk.yellow(`⚠️  Unexpected response: ${dbResponse.status}`));
                }
                
            } catch (dbError) {
                console.log(chalk.red('❌ Failed to connect to Databricks'));
                console.log(chalk.gray(`   Error: ${dbError.message}`));
            }
            
            // Step 3: Display JDBC connection info
            console.log(chalk.blue.bold('\n📋 JDBC Connection Parameters:'));
            console.log(chalk.gray('Use these values in your Java application:\n'));
            
            console.log(chalk.white('Properties for JDBC connection:'));
            console.log(chalk.cyan('  AuthMech: "11"'));
            console.log(chalk.cyan('  Auth_Flow: "1"'));
            console.log(chalk.cyan(`  OAuth2ClientId: "${config.clientId}"`));
            console.log(chalk.cyan(`  OAuth2Secret: "${config.clientSecret}"`));
            console.log(chalk.cyan(`  OAuth2TokenEndpoint: "${tokenEndpoint}"`));
            console.log(chalk.cyan(`  OAuth2Scope: "${config.scope}"`));
            console.log(chalk.cyan('  SSL: "1"'));
            
            return true;
        }
        
    } catch (error) {
        console.log(chalk.red('\n❌ Authentication failed!'));
        
        if (error.response?.data) {
            const errorData = error.response.data;
            console.log(chalk.red(`   Error: ${errorData.error}`));
            console.log(chalk.red(`   Description: ${errorData.error_description}`));
            
            // Common error explanations
            if (errorData.error === 'invalid_client') {
                console.log(chalk.yellow('\n💡 Possible issues:'));
                console.log('   • Client ID is incorrect');
                console.log('   • Client Secret is incorrect or expired');
                console.log('   • Service Principal doesn\'t exist in this tenant');
            } else if (errorData.error === 'invalid_scope') {
                console.log(chalk.yellow('\n💡 The scope may be incorrect for your Databricks instance'));
            } else if (errorData.error === 'unauthorized_client') {
                console.log(chalk.yellow('\n💡 The Service Principal may not have the required permissions'));
            }
        } else {
            console.log(chalk.red(`   ${error.message}`));
            
            if (error.code === 'ENOTFOUND') {
                console.log(chalk.yellow('\n💡 Check your internet connection and tenant ID'));
            }
        }
        
        console.log(chalk.yellow('\n📝 Troubleshooting steps:'));
        console.log('1. Verify Client ID and Client Secret in Azure Portal');
        console.log('2. Ensure the Client Secret hasn\'t expired');
        console.log('3. Check that the Tenant ID is correct');
        console.log('4. Confirm the Service Principal exists in Azure AD');
        
        return false;
    }
}

// Additional utility function to decode JWT token (for debugging)
function decodeToken(token) {
    try {
        const parts = token.split('.');
        if (parts.length !== 3) {
            return null;
        }
        
        const payload = Buffer.from(parts[1], 'base64').toString('utf8');
        return JSON.parse(payload);
    } catch (error) {
        return null;
    }
}

// Test connectivity to Databricks
async function testDatabricksConnectivity() {
    console.log(chalk.yellow('\n3. Testing basic Databricks connectivity...'));
    
    try {
        const response = await axios.get(
            `${config.databricksHost}/api/2.0/clusters/spark-versions`,
            {
                timeout: 5000,
                validateStatus: () => true
            }
        );
        
        if (response.status === 401) {
            console.log(chalk.green('✅ Databricks workspace is reachable'));
            console.log(chalk.gray('   (Authentication required as expected)'));
        } else if (response.status === 200) {
            console.log(chalk.green('✅ Databricks workspace is reachable (public endpoint)'));
        } else {
            console.log(chalk.yellow(`⚠️  Unexpected response: ${response.status}`));
        }
        
        return true;
    } catch (error) {
        console.log(chalk.red('❌ Cannot reach Databricks workspace'));
        console.log(chalk.gray(`   Error: ${error.message}`));
        console.log(chalk.yellow('\n💡 Check:'));
        console.log('   • Workspace URL is correct');
        console.log('   • Network/firewall allows connection');
        console.log('   • VPN connection (if required)');
        return false;
    }
}

// Main execution
async function main() {
    // First test Databricks connectivity
    await testDatabricksConnectivity();
    
    // Then test authentication
    const success = await testAuthentication();
    
    if (success) {
        console.log(chalk.green.bold('\n✅ Authentication test completed successfully!'));
        console.log(chalk.gray('Your Service Principal is ready for JDBC connections.'));
    } else {
        console.log(chalk.red.bold('\n❌ Authentication test failed'));
        console.log(chalk.gray('Please fix the issues above before using JDBC.'));
    }
    
    process.exit(success ? 0 : 1);
}

// Run if executed directly
if (require.main === module) {
    main().catch(error => {
        console.error(chalk.red('Unexpected error:'), error);
        process.exit(1);
    });
}

module.exports = { testAuthentication, testDatabricksConnectivity };