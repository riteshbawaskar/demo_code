// verify-databricks-spn.js
const axios = require('axios');
const https = require('https');

// Configuration
const config = {
    // Databricks workspace configuration
    databricksHost: 'https://your-workspace.cloud.databricks.com',
    databricksToken: 'your-personal-access-token', // PAT for API access
    
    // Service Principal details to verify
    servicePrincipalAppId: 'your-service-principal-application-id',
    servicePrincipalClientId: 'your-service-principal-client-id',
    servicePrincipalSecret: 'your-service-principal-secret',
    
    // Azure AD configuration (for Azure Databricks)
    tenantId: 'your-azure-tenant-id',
    
    // Optional: specific resources to check
    warehouseId: 'your-warehouse-id', // Optional
    clusterId: 'your-cluster-id', // Optional
};

// Create axios instance with default headers
const databricksApi = axios.create({
    baseURL: config.databricksHost,
    headers: {
        'Authorization': `Bearer ${config.databricksToken}`,
        'Content-Type': 'application/json'
    },
    httpsAgent: new https.Agent({
        rejectUnauthorized: true
    })
});

/**
 * Main class for verifying Service Principal configuration
 */
class DatabricksServicePrincipalVerifier {
    constructor(config) {
        this.config = config;
        this.spnDetails = null;
    }

    /**
     * Run all verification checks
     */
    async runFullVerification() {
        console.log('═══════════════════════════════════════════════════════');
        console.log('  DATABRICKS SERVICE PRINCIPAL VERIFICATION');
        console.log('═══════════════════════════════════════════════════════');
        console.log(`Workspace: ${this.config.databricksHost}`);
        console.log(`SPN App ID: ${this.config.servicePrincipalAppId}`);
        console.log('═══════════════════════════════════════════════════════\n');

        try {
            // Step 1: Check if SPN exists
            const spnExists = await this.checkServicePrincipalExists();
            if (!spnExists) {
                console.log('\n❌ VERIFICATION FAILED: Service Principal not found in Databricks!');
                this.printAddInstructions();
                return false;
            }

            // Step 2: Check group memberships
            await this.checkGroupMemberships();

            // Step 3: Check entitlements
            await this.checkEntitlements();

            // Step 4: Check SQL Warehouse permissions
            if (this.config.warehouseId) {
                await this.checkSQLWarehousePermissions();
            }

            // Step 5: Check Cluster permissions
            if (this.config.clusterId) {
                await this.checkClusterPermissions();
            }

            // Step 6: Test authentication
            await this.testServicePrincipalAuthentication();

            // Print summary
            this.printVerificationSummary();
            
            return true;

        } catch (error) {
            console.error('\n❌ Verification failed with error:', error.message);
            return false;
        }
    }

    /**
     * Check if Service Principal exists in Databricks
     */
    async checkServicePrincipalExists() {
        console.log('1. Checking if Service Principal exists...');
        console.log('─────────────────────────────────────────');
        
        try {
            const response = await databricksApi.get('/api/2.0/preview/scim/v2/ServicePrincipals');
            const servicePrincipals = response.data.Resources || [];
            
            // Find our Service Principal
            const ourSPN = servicePrincipals.find(sp => 
                sp.applicationId === this.config.servicePrincipalAppId
            );
            
            if (ourSPN) {
                this.spnDetails = ourSPN;
                console.log(`✅ Service Principal found!`);
                console.log(`   Display Name: ${ourSPN.displayName}`);
                console.log(`   Application ID: ${ourSPN.applicationId}`);
                console.log(`   Internal ID: ${ourSPN.id}`);
                console.log(`   Active: ${ourSPN.active}`);
                return true;
            } else {
                console.log(`❌ Service Principal with App ID ${this.config.servicePrincipalAppId} not found`);
                return false;
            }
        } catch (error) {
            console.error('❌ Error checking Service Principal:', error.response?.data || error.message);
            return false;
        }
    }

    /**
     * Check group memberships
     */
    async checkGroupMemberships() {
        console.log('\n2. Checking Group Memberships...');
        console.log('─────────────────────────────────────────');
        
        try {
            const response = await databricksApi.get('/api/2.0/preview/scim/v2/Groups');
            const groups = response.data.Resources || [];
            const memberships = [];
            
            for (const group of groups) {
                const members = group.members || [];
                const isMember = members.some(member => 
                    member.value === this.spnDetails.id
                );
                
                if (isMember) {
                    memberships.push(group.displayName);
                    console.log(`✅ Member of group: ${group.displayName}`);
                }
            }
            
            if (memberships.length === 0) {
                console.log('⚠️  Service Principal is not a member of any groups');
                console.log('   Recommendation: Add to "users" group for basic access');
            }
            
            this.spnDetails.groupMemberships = memberships;
            return memberships;
            
        } catch (error) {
            console.error('❌ Error checking groups:', error.response?.data || error.message);
            return [];
        }
    }

    /**
     * Check workspace entitlements
     */
    async checkEntitlements() {
        console.log('\n3. Checking Workspace Entitlements...');
        console.log('─────────────────────────────────────────');
        
        try {
            const response = await databricksApi.get(
                `/api/2.0/preview/scim/v2/ServicePrincipals/${this.spnDetails.id}`
            );
            
            const entitlements = response.data.entitlements || [];
            
            if (entitlements.length > 0) {
                entitlements.forEach(ent => {
                    console.log(`✅ Entitlement: ${ent.value}`);
                });
            } else {
                console.log('⚠️  No special entitlements assigned');
                console.log('   Recommendation: Add "workspace-access" entitlement');
            }
            
            this.spnDetails.entitlements = entitlements;
            return entitlements;
            
        } catch (error) {
            console.error('❌ Error checking entitlements:', error.response?.data || error.message);
            return [];
        }
    }

    /**
     * Check SQL Warehouse permissions
     */
    async checkSQLWarehousePermissions() {
        console.log('\n4. Checking SQL Warehouse Permissions...');
        console.log('─────────────────────────────────────────');
        
        try {
            const response = await databricksApi.get(
                `/api/2.0/permissions/warehouses/${this.config.warehouseId}`
            );
            
            const acls = response.data.access_control_list || [];
            const spnPermissions = acls.find(acl => 
                acl.service_principal_name === this.config.servicePrincipalAppId
            );
            
            if (spnPermissions) {
                console.log('✅ SQL Warehouse Permissions found:');
                spnPermissions.all_permissions.forEach(perm => {
                    console.log(`   - ${perm.permission_level}`);
                });
                return true;
            } else {
                console.log('⚠️  No explicit SQL Warehouse permissions found');
                console.log('   Recommendation: Grant "CAN_USE" permission on the warehouse');
                return false;
            }
            
        } catch (error) {
            if (error.response?.status === 404) {
                console.log('ℹ️  SQL Warehouse not found or not accessible');
            } else {
                console.error('❌ Error checking warehouse permissions:', error.response?.data || error.message);
            }
            return false;
        }
    }

    /**
     * Check Cluster permissions
     */
    async checkClusterPermissions() {
        console.log('\n5. Checking Cluster Permissions...');
        console.log('─────────────────────────────────────────');
        
        try {
            const response = await databricksApi.get(
                `/api/2.0/permissions/clusters/${this.config.clusterId}`
            );
            
            const acls = response.data.access_control_list || [];
            const spnPermissions = acls.find(acl => 
                acl.service_principal_name === this.config.servicePrincipalAppId
            );
            
            if (spnPermissions) {
                console.log('✅ Cluster Permissions found:');
                spnPermissions.all_permissions.forEach(perm => {
                    console.log(`   - ${perm.permission_level}`);
                });
                return true;
            } else {
                console.log('⚠️  No explicit Cluster permissions found');
                console.log('   Recommendation: Grant "CAN_ATTACH_TO" permission on the cluster');
                return false;
            }
            
        } catch (error) {
            if (error.response?.status === 404) {
                console.log('ℹ️  Cluster not found or not accessible');
            } else {
                console.error('❌ Error checking cluster permissions:', error.response?.data || error.message);
            }
            return false;
        }
    }

    /**
     * Test Service Principal authentication with OAuth
     */
    async testServicePrincipalAuthentication() {
        console.log('\n6. Testing Service Principal Authentication...');
        console.log('─────────────────────────────────────────');
        
        try {
            // Get Azure AD token
            const tokenEndpoint = `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`;
            
            const tokenResponse = await axios.post(
                tokenEndpoint,
                new URLSearchParams({
                    'grant_type': 'client_credentials',
                    'client_id': this.config.servicePrincipalClientId,
                    'client_secret': this.config.servicePrincipalSecret,
                    'scope': '2ff814a6-3304-4ab8-85cb-cd0e6f879c1d/.default'
                }),
                {
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded'
                    }
                }
            );
            
            if (tokenResponse.data.access_token) {
                console.log('✅ OAuth token obtained successfully');
                console.log('   Service Principal can authenticate with Azure AD');
                
                // Optional: Test Databricks API with SPN token
                await this.testDatabricksAPIWithSPNToken(tokenResponse.data.access_token);
                
                return true;
            }
            
        } catch (error) {
            console.log('❌ Authentication test failed');
            if (error.response?.data) {
                console.log('   Error:', error.response.data.error_description || error.response.data.error);
            } else {
                console.log('   Error:', error.message);
            }
            console.log('   Check: Client ID, Client Secret, and Tenant ID');
            return false;
        }
    }

    /**
     * Test Databricks API with SPN token
     */
    async testDatabricksAPIWithSPNToken(token) {
        try {
            // Try to access Databricks API with SPN token
            const response = await axios.get(
                `${this.config.databricksHost}/api/2.0/clusters/list`,
                {
                    headers: {
                        'Authorization': `Bearer ${token}`
                    }
                }
            );
            
            console.log('✅ Successfully accessed Databricks API with SPN token');
            return true;
            
        } catch (error) {
            if (error.response?.status === 403) {
                console.log('⚠️  SPN token valid but lacks Databricks permissions');
            } else {
                console.log('ℹ️  Could not test Databricks API with SPN token');
            }
            return false;
        }
    }

    /**
     * Print instructions for adding Service Principal
     */
    printAddInstructions() {
        console.log('\n📝 How to add Service Principal to Databricks:');
        console.log('─────────────────────────────────────────────');
        console.log('1. Log into Databricks workspace as admin');
        console.log('2. Go to Admin Settings → Service Principals');
        console.log('3. Click "Add Service Principal"');
        console.log(`4. Enter Application ID: ${this.config.servicePrincipalAppId}`);
        console.log('5. Click "Add"');
        console.log('6. Grant necessary permissions (groups, entitlements, resources)');
    }

    /**
     * Print verification summary
     */
    printVerificationSummary() {
        console.log('\n═══════════════════════════════════════════════════════');
        console.log('  VERIFICATION SUMMARY');
        console.log('═══════════════════════════════════════════════════════');
        
        const checks = [
            {
                name: 'Service Principal exists',
                status: !!this.spnDetails,
                detail: this.spnDetails?.displayName
            },
            {
                name: 'Group memberships',
                status: this.spnDetails?.groupMemberships?.length > 0,
                detail: `${this.spnDetails?.groupMemberships?.length || 0} groups`
            },
            {
                name: 'Entitlements',
                status: this.spnDetails?.entitlements?.length > 0,
                detail: `${this.spnDetails?.entitlements?.length || 0} entitlements`
            }
        ];
        
        checks.forEach(check => {
            const icon = check.status ? '✅' : '⚠️';
            console.log(`${icon} ${check.name}: ${check.detail || 'Not configured'}`);
        });
        
        console.log('\n📋 Required Permissions for JDBC Access:');
        console.log('─────────────────────────────────────────');
        console.log('• Workspace access (via group membership or entitlement)');
        console.log('• SQL Warehouse: "CAN_USE" permission');
        console.log('• Cluster: "CAN_ATTACH_TO" permission');
        console.log('• Tables/Schemas: SELECT, USE SCHEMA grants');
    }
}

/**
 * Utility function to grant permissions via SQL
 */
async function grantSQLPermissions(config) {
    console.log('\n📝 SQL Commands to Grant Permissions:');
    console.log('─────────────────────────────────────────');
    console.log('Run these commands in Databricks SQL:');
    console.log(`
-- Grant catalog access (Unity Catalog)
GRANT USE CATALOG ON CATALOG your_catalog TO \`${config.servicePrincipalAppId}\`;

-- Grant schema usage
GRANT USAGE ON SCHEMA your_schema TO \`${config.servicePrincipalAppId}\`;

-- Grant table access
GRANT SELECT ON TABLE your_schema.your_table TO \`${config.servicePrincipalAppId}\`;

-- Grant all privileges on a schema
GRANT ALL PRIVILEGES ON SCHEMA your_schema TO \`${config.servicePrincipalAppId}\`;
    `);
}

/**
 * Quick check function
 */
async function quickCheck(config) {
    try {
        const response = await databricksApi.get('/api/2.0/preview/scim/v2/ServicePrincipals');
        const exists = response.data.Resources?.some(sp => 
            sp.applicationId === config.servicePrincipalAppId
        );
        
        if (exists) {
            console.log('✅ Service Principal is configured in Databricks');
        } else {
            console.log('❌ Service Principal NOT found in Databricks');
        }
        
        return exists;
    } catch (error) {
        console.error('❌ Error:', error.message);
        return false;
    }
}

// Main execution
async function main() {
    const verifier = new DatabricksServicePrincipalVerifier(config);
    
    // Run full verification
    await verifier.runFullVerification();
    
    // Print SQL permission commands
    grantSQLPermissions(config);
}

// Export for use as a module
module.exports = {
    DatabricksServicePrincipalVerifier,
    quickCheck,
    grantSQLPermissions
};

// Run if executed directly
if (require.main === module) {
    main().catch(console.error);
}