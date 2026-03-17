// ============================================================
// Distribution Queue Processor — Azure Function (Timer Trigger)
// Processes PM_DistributionQueue items server-side
// ============================================================

@description('Environment name')
param environment string = 'prod'

@description('Azure region')
param location string = resourceGroup().location

@description('SharePoint site URL')
param spSiteUrl string = 'https://mf7m.sharepoint.com/sites/PolicyManager'

@description('Azure AD tenant ID')
param tenantId string

@description('Azure AD client ID (app registration)')
param clientId string

@description('Azure AD client secret')
@secure()
param clientSecret string

// --- Naming ---
var prefix = 'dwx-pm-dist'
var suffix = environment
var uniqueSuffix = uniqueString(resourceGroup().id)

// --- Storage Account ---
resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: '${replace(prefix, '-', '')}st${uniqueSuffix}'
  location: location
  sku: {
    name: 'Standard_LRS'
  }
  kind: 'StorageV2'
  properties: {
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
  }
}

// --- App Service Plan (Consumption Y1) ---
resource appServicePlan 'Microsoft.Web/serverfarms@2023-01-01' = {
  name: '${prefix}-plan-${suffix}'
  location: location
  sku: {
    name: 'Y1'
    tier: 'Dynamic'
  }
  kind: 'functionapp'
}

// --- Application Insights ---
resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: '${prefix}-insights-${suffix}'
  location: location
  kind: 'web'
  properties: {
    Application_Type: 'web'
    Request_Source: 'rest'
  }
}

// --- Function App ---
resource functionApp 'Microsoft.Web/sites@2023-01-01' = {
  name: '${prefix}-func-${suffix}'
  location: location
  kind: 'functionapp'
  properties: {
    serverFarmId: appServicePlan.id
    httpsOnly: true
    siteConfig: {
      appSettings: [
        {
          name: 'AzureWebJobsStorage'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value};EndpointSuffix=core.windows.net'
        }
        {
          name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value};EndpointSuffix=core.windows.net'
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
          value: '~20'
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
          name: 'SP_SITE_URL'
          value: spSiteUrl
        }
        {
          name: 'AZURE_TENANT_ID'
          value: tenantId
        }
        {
          name: 'AZURE_CLIENT_ID'
          value: clientId
        }
        {
          name: 'AZURE_CLIENT_SECRET'
          value: clientSecret
        }
      ]
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
      nodeVersion: '~20'
    }
  }
}

// --- Outputs ---
output functionAppName string = functionApp.name
output functionAppUrl string = 'https://${functionApp.properties.defaultHostName}'
output appInsightsName string = appInsights.name
