// ============================================================================
// DWx Policy Manager — AI Chat Assistant Infrastructure
// ============================================================================
// Provisions: Azure Functions (Node.js 18), Storage, App Insights
// Reuses: Existing Azure OpenAI + Key Vault from quiz-generator deployment
//
// Usage:
//   az deployment group create \
//     --resource-group <rg-name> \
//     --template-file main.bicep \
//     --parameters main.parameters.json
// ============================================================================

targetScope = 'resourceGroup'

// ============================================================================
// Parameters
// ============================================================================

@description('Base name prefix for all resources (e.g., dwx-pm)')
@minLength(3)
@maxLength(15)
param baseName string = 'dwx-pm'

@description('Azure region for all resources')
param location string = resourceGroup().location

@description('Deployment environment')
@allowed(['dev', 'staging', 'prod'])
param environment string = 'prod'

@description('SharePoint site URL for CORS configuration')
param sharePointSiteUrl string = 'https://mf7m.sharepoint.com'

@description('Node.js runtime version for Functions')
param nodeVersion string = '~18'

@description('Existing Azure OpenAI endpoint URL (from quiz-generator deployment)')
param existingOpenAiEndpoint string

@description('Existing Key Vault name (from quiz-generator deployment)')
param existingKeyVaultName string

@description('Resource group containing the existing Key Vault')
param existingKeyVaultResourceGroup string

@description('Azure OpenAI deployment name')
param openAiDeploymentName string = 'gpt-4o'

// ============================================================================
// Variables
// ============================================================================

var uniqueSuffix = uniqueString(resourceGroup().id, baseName, 'chat')
#disable-next-line BCP334
var storageAccountName = toLower(replace('${baseName}chst${uniqueSuffix}', '-', ''))
var functionAppName = '${baseName}-chat-func-${environment}'
var appServicePlanName = '${baseName}-chat-plan-${environment}'
var appInsightsName = '${baseName}-chat-insights-${environment}'
var logAnalyticsName = '${baseName}-chat-logs-${environment}'

var tags = {
  project: 'DWx Policy Manager'
  component: 'AI Chat Assistant'
  environment: environment
  managedBy: 'Bicep'
}

// ============================================================================
// Log Analytics Workspace
// ============================================================================

resource logAnalytics 'Microsoft.OperationalInsights/workspaces@2023-09-01' = {
  name: logAnalyticsName
  location: location
  tags: tags
  properties: {
    sku: { name: 'PerGB2018' }
    retentionInDays: 30
  }
}

// ============================================================================
// Application Insights
// ============================================================================

resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: appInsightsName
  location: location
  tags: tags
  kind: 'web'
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: logAnalytics.id
    RetentionInDays: 30
  }
}

// ============================================================================
// Storage Account (required by Azure Functions)
// ============================================================================

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: take(storageAccountName, 24)
  location: location
  tags: tags
  kind: 'StorageV2'
  sku: { name: 'Standard_LRS' }
  properties: {
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
  }
}

// ============================================================================
// App Service Plan (Consumption / Serverless)
// ============================================================================

resource appServicePlan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: appServicePlanName
  location: location
  tags: tags
  kind: 'functionapp'
  sku: {
    name: 'Y1'
    tier: 'Dynamic'
  }
  properties: { reserved: false }
}

// ============================================================================
// Existing Key Vault Reference
// ============================================================================

// Key Vault lives in a different resource group — RBAC is assigned via module below

// ============================================================================
// Azure Functions App
// ============================================================================

resource functionApp 'Microsoft.Web/sites@2023-12-01' = {
  name: functionAppName
  location: location
  tags: tags
  kind: 'functionapp'
  identity: { type: 'SystemAssigned' }
  properties: {
    serverFarmId: appServicePlan.id
    httpsOnly: true
    siteConfig: {
      nodeVersion: nodeVersion
      cors: {
        allowedOrigins: [
          sharePointSiteUrl
          'https://localhost:4321'
        ]
        supportCredentials: false
      }
      appSettings: [
        {
          name: 'AzureWebJobsStorage'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};EndpointSuffix=${az.environment().suffixes.storage};AccountKey=${storageAccount.listKeys().keys[0].value}'
        }
        {
          name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};EndpointSuffix=${az.environment().suffixes.storage};AccountKey=${storageAccount.listKeys().keys[0].value}'
        }
        {
          name: 'WEBSITE_CONTENTSHARE'
          value: toLower(functionAppName)
        }
        {
          name: 'FUNCTIONS_EXTENSION_VERSION'
          value: '~4'
        }
        {
          name: 'FUNCTIONS_WORKER_RUNTIME'
          value: 'node'
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~18'
        }
        {
          name: 'APPINSIGHTS_INSTRUMENTATIONKEY'
          value: appInsights.properties.InstrumentationKey
        }
        {
          name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
          value: appInsights.properties.ConnectionString
        }
        {
          name: 'AZURE_OPENAI_ENDPOINT'
          value: existingOpenAiEndpoint
        }
        {
          name: 'AZURE_OPENAI_API_KEY'
          value: '@Microsoft.KeyVault(VaultName=${existingKeyVaultName};SecretName=azure-openai-api-key)'
        }
        {
          name: 'AZURE_OPENAI_DEPLOYMENT'
          value: openAiDeploymentName
        }
        {
          name: 'AZURE_OPENAI_API_VERSION'
          value: '2024-02-15-preview'
        }
      ]
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
    }
  }
}

// ============================================================================
// RBAC: Function App → Key Vault Secrets User (cross-resource-group module)
// ============================================================================

module kvRbac 'kvRbac.bicep' = {
  name: 'kv-rbac-${functionAppName}'
  scope: resourceGroup(existingKeyVaultResourceGroup)
  params: {
    keyVaultName: existingKeyVaultName
    principalId: functionApp.identity.principalId
  }
}

// ============================================================================
// Outputs
// ============================================================================

@description('Function App name')
output functionAppName string = functionApp.name

@description('Function App default hostname')
output functionAppUrl string = 'https://${functionApp.properties.defaultHostName}'

@description('Application Insights instrumentation key')
output appInsightsKey string = appInsights.properties.InstrumentationKey

@description('Function App principal ID (for additional RBAC)')
output functionAppPrincipalId string = functionApp.identity.principalId
