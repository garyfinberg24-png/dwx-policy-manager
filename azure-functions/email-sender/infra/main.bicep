// ============================================================================
// DWx Policy Manager — Email Queue Processor (Logic App)
// ============================================================================
// Provisions: Logic App (Consumption), Office 365 + SharePoint API connections
// Reads PM_EmailQueue SharePoint list, sends emails via Office 365, updates status.
//
// Usage:
//   az deployment group create \
//     --resource-group <rg-name> \
//     --template-file main.bicep \
//     --parameters main.parameters.json
//
// Post-deployment: Authorize both API connections in the Azure Portal.
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

@description('SharePoint site URL containing PM_EmailQueue list')
param sharePointSiteUrl string = 'https://mf7m.sharepoint.com/sites/PolicyManager'

@description('SharePoint list name for the email queue')
param emailQueueListName string = 'PM_EmailQueue'

@description('How often (in minutes) the Logic App polls for queued emails')
@minValue(1)
@maxValue(60)
param pollingIntervalMinutes int = 5

@description('Max emails to process per run')
@minValue(1)
@maxValue(50)
param batchSize int = 20

@description('Max send attempts before marking as permanently Failed')
param maxRetryAttempts int = 3

@description('Email address of the shared mailbox to send from. Leave empty to use the connection user mailbox.')
param senderEmailAddress string = ''

// ============================================================================
// Variables
// ============================================================================

var logicAppName = '${baseName}-email-sender-${environment}'
var office365ConnectionName = 'office365-${environment}'
var sharepointConnectionName = 'sharepointonline-${environment}'

var tags = {
  project: 'DWx Policy Manager'
  component: 'Email Sender'
  environment: environment
  managedBy: 'Bicep'
}

// ============================================================================
// API Connections — Office 365 + SharePoint Online
// ============================================================================

resource office365Connection 'Microsoft.Web/connections@2016-06-01' = {
  name: office365ConnectionName
  location: location
  tags: tags
  properties: {
    displayName: 'DWx PM Office 365 Email'
    api: {
      id: subscriptionResourceId('Microsoft.Web/locations/managedApis', location, 'office365')
    }
  }
}

resource sharepointConnection 'Microsoft.Web/connections@2016-06-01' = {
  name: sharepointConnectionName
  location: location
  tags: tags
  properties: {
    displayName: 'DWx PM SharePoint Online'
    api: {
      id: subscriptionResourceId('Microsoft.Web/locations/managedApis', location, 'sharepointonline')
    }
  }
}

// ============================================================================
// Logic App — Email Queue Processor
// ============================================================================

resource logicApp 'Microsoft.Logic/workflows@2019-05-01' = {
  name: logicAppName
  location: location
  tags: tags
  properties: {
    state: 'Enabled'
    definition: {
      '$schema': 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#'
      contentVersion: '1.0.0.0'
      parameters: {
        '$connections': {
          defaultValue: {}
          type: 'Object'
        }
        sharePointSiteUrl: {
          defaultValue: sharePointSiteUrl
          type: 'String'
        }
        emailQueueListName: {
          defaultValue: emailQueueListName
          type: 'String'
        }
        maxRetryAttempts: {
          defaultValue: maxRetryAttempts
          type: 'Int'
        }
        senderEmailAddress: {
          defaultValue: senderEmailAddress
          type: 'String'
        }
      }

      // ── Trigger: Poll every N minutes ──
      triggers: {
        Poll_Email_Queue: {
          type: 'Recurrence'
          recurrence: {
            frequency: 'Minute'
            interval: pollingIntervalMinutes
          }
        }
      }

      // ── Actions ──
      actions: {

        // Step 1: Query PM_EmailQueue for Status='Queued'
        Get_Queued_Emails: {
          type: 'ApiConnection'
          runAfter: {}
          inputs: {
            host: {
              connection: {
                name: '@parameters(\'$connections\')[\'sharepointonline\'][\'connectionId\']'
              }
            }
            method: 'get'
            path: '/datasets/@{encodeURIComponent(encodeURIComponent(\'${sharePointSiteUrl}\'))}/tables/@{encodeURIComponent(encodeURIComponent(\'${emailQueueListName}\'))}/items'
            queries: {
              '$filter': 'Status eq \'Queued\''
              '$orderby': 'Priority desc, QueuedAt asc'
              '$top': batchSize
            }
          }
        }

        // Step 2: Process each email sequentially
        For_Each_Queued_Email: {
          type: 'Foreach'
          runAfter: {
            Get_Queued_Emails: [ 'Succeeded' ]
          }
          foreach: '@body(\'Get_Queued_Emails\')?[\'value\']'
          operationOptions: 'Sequential'
          actions: {

            // 2a: Mark as Processing
            Set_Status_Processing: {
              type: 'ApiConnection'
              runAfter: {}
              inputs: {
                host: {
                  connection: {
                    name: '@parameters(\'$connections\')[\'sharepointonline\'][\'connectionId\']'
                  }
                }
                method: 'patch'
                path: '/datasets/@{encodeURIComponent(encodeURIComponent(\'${sharePointSiteUrl}\'))}/tables/@{encodeURIComponent(encodeURIComponent(\'${emailQueueListName}\'))}/items/@{encodeURIComponent(items(\'For_Each_Queued_Email\')?[\'ID\'])}'
                body: {
                  Status: 'Processing'
                  LastAttemptAt: '@{utcNow()}'
                }
              }
            }

            // 2b: Send via Office 365
            Send_Email: {
              type: 'ApiConnection'
              runAfter: {
                Set_Status_Processing: [ 'Succeeded' ]
              }
              inputs: {
                host: {
                  connection: {
                    name: '@parameters(\'$connections\')[\'office365\'][\'connectionId\']'
                  }
                }
                method: 'post'
                path: '/v2/Mail'
                body: {
                  To: '@{replace(items(\'For_Each_Queued_Email\')?[\'To\'], \';\', \',\')}'
                  Cc: '@{if(empty(items(\'For_Each_Queued_Email\')?[\'CC\']), \'\', replace(items(\'For_Each_Queued_Email\')?[\'CC\'], \';\', \',\'))}'
                  Subject: '@items(\'For_Each_Queued_Email\')?[\'Subject\']'
                  Body: '@items(\'For_Each_Queued_Email\')?[\'Body\']'
                  Importance: '@{if(equals(items(\'For_Each_Queued_Email\')?[\'Priority\'], \'Urgent\'), \'High\', if(equals(items(\'For_Each_Queued_Email\')?[\'Priority\'], \'High\'), \'High\', if(equals(items(\'For_Each_Queued_Email\')?[\'Priority\'], \'Low\'), \'Low\', \'Normal\')))}'
                  IsHtml: true
                  From: '@{if(empty(parameters(\'senderEmailAddress\')), \'\', parameters(\'senderEmailAddress\'))}'
                }
              }
            }

            // 2c: Success → Mark as Sent
            Mark_As_Sent: {
              type: 'ApiConnection'
              runAfter: {
                Send_Email: [ 'Succeeded' ]
              }
              inputs: {
                host: {
                  connection: {
                    name: '@parameters(\'$connections\')[\'sharepointonline\'][\'connectionId\']'
                  }
                }
                method: 'patch'
                path: '/datasets/@{encodeURIComponent(encodeURIComponent(\'${sharePointSiteUrl}\'))}/tables/@{encodeURIComponent(encodeURIComponent(\'${emailQueueListName}\'))}/items/@{encodeURIComponent(items(\'For_Each_Queued_Email\')?[\'ID\'])}'
                body: {
                  Status: 'Sent'
                  SentAt: '@{utcNow()}'
                  AttemptCount: '@{add(int(coalesce(items(\'For_Each_Queued_Email\')?[\'AttemptCount\'], \'0\')), 1)}'
                }
              }
            }

            // 2d: Failure → Retry or mark Failed
            Handle_Send_Failure: {
              type: 'ApiConnection'
              runAfter: {
                Send_Email: [ 'Failed', 'TimedOut' ]
              }
              inputs: {
                host: {
                  connection: {
                    name: '@parameters(\'$connections\')[\'sharepointonline\'][\'connectionId\']'
                  }
                }
                method: 'patch'
                path: '/datasets/@{encodeURIComponent(encodeURIComponent(\'${sharePointSiteUrl}\'))}/tables/@{encodeURIComponent(encodeURIComponent(\'${emailQueueListName}\'))}/items/@{encodeURIComponent(items(\'For_Each_Queued_Email\')?[\'ID\'])}'
                body: {
                  Status: '@{if(greaterOrEquals(add(int(coalesce(items(\'For_Each_Queued_Email\')?[\'AttemptCount\'], \'0\')), 1), parameters(\'maxRetryAttempts\')), \'Failed\', \'Queued\')}'
                  AttemptCount: '@{add(int(coalesce(items(\'For_Each_Queued_Email\')?[\'AttemptCount\'], \'0\')), 1)}'
                  LastAttemptAt: '@{utcNow()}'
                  ErrorMessage: '@{coalesce(body(\'Send_Email\')?[\'error\']?[\'message\'], actions(\'Send_Email\')?[\'error\']?[\'message\'], \'Send failed — see Logic App run history for details\')}'
                }
              }
            }
          }
        }
      }
      outputs: {}
    }
    parameters: {
      '$connections': {
        value: {
          office365: {
            connectionId: office365Connection.id
            connectionName: office365ConnectionName
            id: subscriptionResourceId('Microsoft.Web/locations/managedApis', location, 'office365')
          }
          sharepointonline: {
            connectionId: sharepointConnection.id
            connectionName: sharepointConnectionName
            id: subscriptionResourceId('Microsoft.Web/locations/managedApis', location, 'sharepointonline')
          }
        }
      }
    }
  }
}

// ============================================================================
// Outputs
// ============================================================================

@description('Logic App name')
output logicAppName string = logicApp.name

@description('Logic App resource ID')
output logicAppId string = logicApp.id

@description('Office 365 connection resource ID — authorize in Portal after deployment')
output office365ConnectionId string = office365Connection.id

@description('SharePoint connection resource ID — authorize in Portal after deployment')
output sharepointConnectionId string = sharepointConnection.id
