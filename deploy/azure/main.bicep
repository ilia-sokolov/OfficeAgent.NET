// OfficeAgent.NET MCP server on Azure Container Apps, for Copilot Studio, Microsoft 365 Copilot
// and other hosted clients. Acts as the signed-in user against SharePoint through On-Behalf-Of.
// See README.md. After changing this file, regenerate azuredeploy.json:
//   az bicep build --file deploy/azure/main.bicep --outfile deploy/azure/azuredeploy.json

@description('Region for every resource.')
param location string = resourceGroup().location

@description('Base name for the resources. Lowercase letters, digits and hyphens.')
@minLength(3)
@maxLength(20)
param name string = 'officeagent'

@description('Server image. Pin an exact version; review release notes before moving it.')
param image string = 'ghcr.io/ilia-sokolov/officeagent-mcp:1.2.0'

@description('Entra tenant of the signed-in users.')
param tenantId string = subscription().tenantId

@description('Application (client) id of the middle-tier API app registration.')
param apiClientId string

@description('Client secret of that app registration, used for the On-Behalf-Of exchange.')
@secure()
param apiClientSecret string

@description('Connection id agents address documents under.')
param connectionId string = 'sharepoint'

@description('Instances kept running. 0 saves money, and the first call after idle starts slowly.')
@minValue(0)
@maxValue(1)
param minReplicas int = 1

@description('Author tracked changes and comments with the signed-in user\'s display name.')
param revisionAuthorFromActor bool = false

var suffix = uniqueString(resourceGroup().id, name)
var share = 'officeagent-data'
var issuer = '${environment().authentication.loginEndpoint}${tenantId}/v2.0'

resource logs 'Microsoft.OperationalInsights/workspaces@2023-09-01' = {
  name: '${name}-logs-${suffix}'
  location: location
  properties: {
    sku: { name: 'PerGB2018' }
    retentionInDays: 30
  }
}

resource storage 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: take('oa${suffix}', 24)
  location: location
  sku: { name: 'Standard_LRS' }
  kind: 'StorageV2'
  properties: {
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
    supportsHttpsTrafficOnly: true
  }
}

resource files 'Microsoft.Storage/storageAccounts/fileServices@2023-05-01' = {
  parent: storage
  name: 'default'
}

resource dataShare 'Microsoft.Storage/storageAccounts/fileServices/shares@2023-05-01' = {
  parent: files
  name: share
  properties: { shareQuota: 1 }
}

resource managedEnvironment 'Microsoft.App/managedEnvironments@2024-03-01' = {
  name: '${name}-env-${suffix}'
  location: location
  properties: {
    appLogsConfiguration: {
      destination: 'log-analytics'
      logAnalyticsConfiguration: {
        customerId: logs.properties.customerId
        sharedKey: logs.listKeys().primarySharedKey
      }
    }
  }
}

resource environmentStorage 'Microsoft.App/managedEnvironments/storages@2024-03-01' = {
  parent: managedEnvironment
  name: 'data'
  properties: {
    azureFile: {
      accountName: storage.name
      accountKey: storage.listKeys().keys[0].value
      shareName: dataShare.name
      accessMode: 'ReadWrite'
    }
  }
}

resource app 'Microsoft.App/containerApps@2024-03-01' = {
  name: name
  location: location
  properties: {
    managedEnvironmentId: managedEnvironment.id
    configuration: {
      ingress: {
        external: true
        targetPort: 8080
        allowInsecure: false
      }
      secrets: [
        { name: 'api-client-secret', value: apiClientSecret }
      ]
    }
    template: {
      containers: [
        {
          name: 'officeagent-mcp'
          image: image
          resources: { cpu: json('0.5'), memory: '1Gi' }
          env: [
            // Copilot Studio asks the user for every required input the model leaves out.
            { name: 'OfficeAgent__StrictToolSchemas', value: 'false' }
            // A retried call must not apply the same plan twice.
            { name: 'OfficeAgent__RequirePlanSnapshot', value: 'true' }
            { name: 'OfficeAgent__RevisionAuthorFromActor', value: string(revisionAuthorFromActor) }
            { name: 'OfficeAgent__SharePointConnections__0__ConnectionId', value: connectionId }
            { name: 'OfficeAgent__SharePointConnections__0__AuthMode', value: 'onBehalfOf' }
            { name: 'OfficeAgent__SharePointConnections__0__TenantId', value: tenantId }
            { name: 'OfficeAgent__SharePointConnections__0__ClientId', value: apiClientId }
            { name: 'OfficeAgent__SharePointConnections__0__ClientSecret', secretRef: 'api-client-secret' }
            { name: 'OfficeAgent__SharePointConnections__0__RegistrationIndexPath', value: '/data/registrations.json' }
            { name: 'OfficeAgent__SharePointConnections__0__AllowedExtensions__0', value: '.docx' }
            { name: 'OfficeAgent__SharePointConnections__0__AllowedExtensions__1', value: '.pptx' }
            { name: 'OfficeAgent__SharePointConnections__0__AllowedExtensions__2', value: '.xlsx' }
          ]
          volumeMounts: [
            { volumeName: 'data', mountPath: '/data' }
          ]
          probes: [
            {
              type: 'Liveness'
              httpGet: { path: '/healthz', port: 8080 }
              initialDelaySeconds: 10
              periodSeconds: 30
            }
          ]
        }
      ]
      // One writer: the registration index is a single JSON file.
      scale: {
        minReplicas: minReplicas
        maxReplicas: 1
      }
      volumes: [
        { name: 'data', storageType: 'AzureFile', storageName: environmentStorage.name }
      ]
    }
  }
}

// The server has no authentication of its own. Container Apps validates the caller's Entra token
// here and forwards it unchanged, and the server exchanges it for the user's Graph access.
resource auth 'Microsoft.App/containerApps/authConfigs@2024-03-01' = {
  parent: app
  name: 'current'
  properties: {
    platform: { enabled: true }
    globalValidation: {
      unauthenticatedClientAction: 'Return401'
      excludedPaths: [ '/healthz' ]
    }
    identityProviders: {
      azureActiveDirectory: {
        enabled: true
        registration: {
          clientId: apiClientId
          openIdIssuer: issuer
        }
        validation: {
          allowedAudiences: [ apiClientId, 'api://${apiClientId}' ]
        }
      }
    }
  }
}

@description('MCP endpoint to give a client.')
output serverUrl string = 'https://${app.properties.configuration.ingress.fqdn}/'
