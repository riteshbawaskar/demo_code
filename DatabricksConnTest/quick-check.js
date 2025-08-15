// quick-check.js - Simplified Service Principal verification
require('dotenv').config();
const axios = require('axios');
const chalk = require('chalk');

// Load configuration from environment variables
const config = {
    databricksHost: process.env.DATABRICKS_HOST || 'https://your-workspace.cloud.databricks.com',
    databricksToken: process.env.DATABRICKS_TOKEN || 'your-personal-access-token',
    servicePrincipalAppId: process.env.SERVICE_PRINCIPAL_APP_ID || 'your-app-id',
};

/**
 * Quick check if Service Principal exists and has basic permissions
 */
async function checkServicePrincipal() {
    console.log(chalk.blue.bold('\n🔍 Databricks Service Principal Quick Check\n'));
    console.log(chalk.gray(`Workspace: ${config.databricksHost}`));
    console.log(chalk.gray(`SPN App ID: ${config.servicePrincipalAppId}\n`));

    const api = axios.create({
        baseURL: config.databricksHost,
        headers: {
            'Authorization': `Bearer ${config.databricksToken}`,
            'Content-Type': 'application/json'
        }
    });

    try {
        // Step 1: Check if SPN exists
        console.log(chalk.yellow('Checking Service Principal...'));
        const spnResponse = await api.get('/api/2.0/preview/scim/v2/ServicePrincipals');
        const spn = spnResponse.data.Resources?.find(sp => 
            sp.applicationId === config.servicePrincipalAppId
        );

        if (!spn) {
            console.log(chalk.red('❌ Service Principal NOT found in Databricks!'));
            console.log(chalk.yellow('\nTo add it:'));
            console.log('1. Go to Admin Settings → Service Principals');
            console.log('2. Click "Add Service Principal"');
            console.log(`3. Enter Application ID: ${config.servicePrincipalAppId}`);
            return false;
        }

        console.log(chalk.green(`✅ Service Principal found: ${spn.displayName}`));
        console.log(chalk.gray(`   ID: ${spn.id}`));
        console.log(chalk.gray(`   Active: ${spn.active}`));

        // Step 2: Check groups
        console.log(chalk.yellow('\nChecking group memberships...'));
        const groupsResponse = await api.get('/api/2.0/preview/scim/v2/Groups');
        const groups = groupsResponse.data.Resources || [];
        const memberships = [];

        for (const group of groups) {
            if (group.members?.some(m => m.value === spn.id)) {
                memberships.push(group.displayName);
            }
        }

        if (memberships.length > 0) {
            console.log(chalk.green(`✅ Member of ${memberships.length} group(s):`));
            memberships.forEach(g => console.log(chalk.gray(`   - ${g}`)));
        } else {
            console.log(chalk.yellow('⚠️  Not a member of any groups'));
        }

        // Step 3: Check entitlements
        console.log(chalk.yellow('\nChecking entitlements...'));
        const spnDetailsResponse = await api.get(
            `/api/2.0/preview/scim/v2/ServicePrincipals/${spn.id}`
        );
        const entitlements = spnDetailsResponse.data.entitlements || [];

        if (entitlements.length > 0) {
            console.log(chalk.green(`✅ Has ${entitlements.length} entitlement(s):`));
            entitlements.forEach(e => console.log(chalk.gray(`   - ${e.value}`)));
        } else {
            console.log(chalk.yellow('⚠️  No entitlements assigned'));
        }

        // Summary
        console.log(chalk.blue.bold('\n📊 Summary:'));
        const hasBasicAccess = memberships.length > 0 || entitlements.length > 0;
        
        if (hasBasicAccess) {
            console.log(chalk.green('✅ Service Principal is configured and has basic access'));
            console.log(chalk.gray('\nNext steps:'));
            console.log('• Grant specific SQL Warehouse or Cluster permissions as needed');
            console.log('• Test JDBC connection with your application');
        } else {
            console.log(chalk.yellow('⚠️  Service Principal exists but may lack permissions'));
            console.log(chalk.gray('\nRecommended actions:'));
            console.log('• Add to "users" group for basic workspace access');
            console.log('• Grant "workspace-access" entitlement');
            console.log('• Set permissions on specific resources (warehouses/clusters)');
        }

        return true;

    } catch (error) {
        console.log(chalk.red('\n❌ Error during verification:'));
        if (error.response) {
            console.log(chalk.red(`   Status: ${error.response.status}`));
            console.log(chalk.red(`   Message: ${error.response.data?.message || error.message}`));
        } else {
            console.log(chalk.red(`   ${error.message}`));
        }
        
        if (error.response?.status === 401) {
            console.log(chalk.yellow('\n💡 Check your Databricks personal access token'));
        }
        
        return false;
    }
}

// Run the check
checkServicePrincipal()
    .then(success => {
        process.exit(success ? 0 : 1);
    })
    .catch(error => {
        console.error(chalk.red('Unexpected error:'), error);
        process.exit(1);
    });