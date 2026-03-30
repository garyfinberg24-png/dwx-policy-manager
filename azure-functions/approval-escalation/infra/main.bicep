// ============================================================================
// DWx Policy Manager — Approval Escalation (Logic App)
// ============================================================================
// Provisions: Logic App (Consumption), SharePoint API connection
// Polls PM_Approvals SharePoint list for overdue approvals, marks them as
// escalated, and queues urgent notifications to PM_NotificationQueue.
//
// Usage:
//   az deployment group create \
//     --resource-group <rg-name> \
//     --template-file main.bicep \
//     --parameters main.parameters.json
//
// Post-deployment: Authorize the SharePoint API connection in the Azure Portal.
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

@description('SharePoint site URL containing PM_Approvals and PM_NotificationQueue lists')
param sharePointSiteUrl string = 'https://mf7m.sharepoint.com/sites/PolicyManager'

@description('How often (in minutes) the Logic App polls for overdue approvals')
@minValue(1)
@maxValue(60)
param pollingIntervalMinutes int = 15

@description('SharePoint list name for approvals')
param escalationListName string = 'PM_Approvals'

@description('SharePoint list name for the notification queue')
param notificationQueueListName string = 'PM_NotificationQueue'

// ============================================================================
// Variables
// ============================================================================

var logicAppName = '${baseName}-approval-escalation-${environment}'
var sharepointConnectionName = 'sharepointonline-${environment}'

var tags = {
  project: 'DWx Policy Manager'
  component: 'Approval Escalation'
  environment: environment
  managedBy: 'Bicep'
}

// ============================================================================
// API Connection — SharePoint Online
// ============================================================================

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
// Logic App — Approval Escalation Processor
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
        escalationListName: {
          defaultValue: escalationListName
          type: 'String'
        }
        notificationQueueListName: {
          defaultValue: notificationQueueListName
          type: 'String'
        }
      }

      // ── Trigger: Poll every N minutes ──
      triggers: {
        Poll_Overdue_Approvals: {
          type: 'Recurrence'
          recurrence: {
            frequency: 'Minute'
            interval: pollingIntervalMinutes
          }
        }
      }

      // ── Actions ──
      actions: {

        // Step 1: Query PM_Approvals for Status='Pending' AND DueDate is past
        Get_Overdue_Approvals: {
          type: 'ApiConnection'
          runAfter: {}
          inputs: {
            host: {
              connection: {
                name: '@parameters(\'$connections\')[\'sharepointonline\'][\'connectionId\']'
              }
            }
            method: 'get'
            path: '/datasets/@{encodeURIComponent(encodeURIComponent(\'${sharePointSiteUrl}\'))}/tables/@{encodeURIComponent(encodeURIComponent(\'${escalationListName}\'))}/items'
            queries: {
              '$filter': 'Status eq \'Pending\' and DueDate lt \'@{utcNow()}\''
              '$orderby': 'DueDate asc'
            }
          }
        }

        // Step 2: Process each overdue approval sequentially
        For_Each_Overdue_Approval: {
          type: 'Foreach'
          runAfter: {
            Get_Overdue_Approvals: [ 'Succeeded' ]
          }
          foreach: '@body(\'Get_Overdue_Approvals\')?[\'value\']'
          operationOptions: 'Sequential'
          actions: {

            // 2a: Update approval — mark as overdue and escalated
            Update_Approval_Status: {
              type: 'ApiConnection'
              runAfter: {}
              inputs: {
                host: {
                  connection: {
                    name: '@parameters(\'$connections\')[\'sharepointonline\'][\'connectionId\']'
                  }
                }
                method: 'patch'
                path: '/datasets/@{encodeURIComponent(encodeURIComponent(\'${sharePointSiteUrl}\'))}/tables/@{encodeURIComponent(encodeURIComponent(\'${escalationListName}\'))}/items/@{encodeURIComponent(items(\'For_Each_Overdue_Approval\')?[\'ID\'])}'
                body: {
                  IsOverdue: true
                  Status: 'Escalated'
                }
              }
            }

            // 2b: Queue urgent notification to PM_NotificationQueue
            Create_Notification_Queue_Item: {
              type: 'ApiConnection'
              runAfter: {
                Update_Approval_Status: [ 'Succeeded' ]
              }
              inputs: {
                host: {
                  connection: {
                    name: '@parameters(\'$connections\')[\'sharepointonline\'][\'connectionId\']'
                  }
                }
                method: 'post'
                path: '/datasets/@{encodeURIComponent(encodeURIComponent(\'${sharePointSiteUrl}\'))}/tables/@{encodeURIComponent(encodeURIComponent(\'${notificationQueueListName}\'))}/items'
                body: {
                  Title: 'Approval Overdue'
                  RecipientEmail: '@{items(\'For_Each_Overdue_Approval\')?[\'ApproverEmail\']}'
                  Priority: 'Urgent'
                  Message: 'The approval for "@{items(\'For_Each_Overdue_Approval\')?[\'Title\']}" was due on @{items(\'For_Each_Overdue_Approval\')?[\'DueDate\']} and has been escalated. Please review immediately.'
                  QueueStatus: 'Pending'
                  Channel: 'Email'
                  QueuedAt: '@{utcNow()}'
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

@description('SharePoint connection resource ID — authorize in Portal after deployment')
output apiConnectionId string = sharepointConnection.id
