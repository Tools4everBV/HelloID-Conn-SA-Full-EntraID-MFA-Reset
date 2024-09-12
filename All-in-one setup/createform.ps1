# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

#HelloID variables
#Note: when running this script inside HelloID; portalUrl and API credentials are provided automatically (generate and save API credentials first in your admin panel!)
$portalUrl = "https://CUSTOMER.helloid.com"
$apiKey = "API_KEY"
$apiSecret = "API_SECRET"
$delegatedFormAccessGroupNames = @("") #Only unique names are supported. Groups must exist!
$delegatedFormCategories = @("Entra ID","User Management") #Only unique names are supported. Categories will be created if not exists
$script:debugLogging = $false #Default value: $false. If $true, the HelloID resource GUIDs will be shown in the logging
$script:duplicateForm = $false #Default value: $false. If $true, the HelloID resource names will be changed to import a duplicate Form
$script:duplicateFormSuffix = "_tmp" #the suffix will be added to all HelloID resource names to generate a duplicate form with different resource names

#The following HelloID Global variables are used by this form. No existing HelloID global variables will be overriden only new ones are created.
#NOTE: You can also update the HelloID Global variable values afterwards in the HelloID Admin Portal: https://<CUSTOMER>.helloid.com/admin/variablelibrary
$globalHelloIDVariables = [System.Collections.Generic.List[object]]@();

#Global variable #1 >> EntraAppSecret
$tmpName = @'
EntraAppSecret
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True"});

#Global variable #2 >> EntraTenantId
$tmpName = @'
EntraTenantId
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #3 >> EntraAppId
$tmpName = @'
EntraAppId
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#make sure write-information logging is visual
$InformationPreference = "continue"

# Check for prefilled API Authorization header
if (-not [string]::IsNullOrEmpty($portalApiBasic)) {
    $script:headers = @{"authorization" = $portalApiBasic}
    Write-Information "Using prefilled API credentials"
} else {
    # Create authorization headers with HelloID API key
    $pair = "$apiKey" + ":" + "$apiSecret"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $key = "Basic $base64"
    $script:headers = @{"authorization" = $Key}
    Write-Information "Using manual API credentials"
}

# Check for prefilled PortalBaseURL
if (-not [string]::IsNullOrEmpty($portalBaseUrl)) {
    $script:PortalBaseUrl = $portalBaseUrl
    Write-Information "Using prefilled PortalURL: $script:PortalBaseUrl"
} else {
    $script:PortalBaseUrl = $portalUrl
    Write-Information "Using manual PortalURL: $script:PortalBaseUrl"
}

# Define specific endpoint URI
$script:PortalBaseUrl = $script:PortalBaseUrl.trim("/") + "/"  

# Make sure to reveive an empty array using PowerShell Core
function ConvertFrom-Json-WithEmptyArray([string]$jsonString) {
    # Running in PowerShell Core?
    if($IsCoreCLR -eq $true){
        $r = [Object[]]($jsonString | ConvertFrom-Json -NoEnumerate)
        return ,$r  # Force return value to be an array using a comma
    } else {
        $r = [Object[]]($jsonString | ConvertFrom-Json)
        return ,$r  # Force return value to be an array using a comma
    }
}

function Invoke-HelloIDGlobalVariable {
    param(
        [parameter(Mandatory)][String]$Name,
        [parameter(Mandatory)][String][AllowEmptyString()]$Value,
        [parameter(Mandatory)][String]$Secret
    )

    $Name = $Name + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl + "api/v1/automation/variables/named/$Name")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
    
        if ([string]::IsNullOrEmpty($response.automationVariableGuid)) {
            #Create Variable
            $body = @{
                name     = $Name;
                value    = $Value;
                secret   = $Secret;
                ItemType = 0;
            }    
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl + "api/v1/automation/variable")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $variableGuid = $response.automationVariableGuid

            Write-Information "Variable '$Name' created$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        } else {
            $variableGuid = $response.automationVariableGuid
            Write-Warning "Variable '$Name' already exists$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        }
    } catch {
        Write-Error "Variable '$Name', message: $_"
    }
}

function Invoke-HelloIDAutomationTask {
    param(
        [parameter(Mandatory)][String]$TaskName,
        [parameter(Mandatory)][String]$UseTemplate,
        [parameter(Mandatory)][String]$AutomationContainer,
        [parameter(Mandatory)][String][AllowEmptyString()]$Variables,
        [parameter(Mandatory)][String]$PowershellScript,
        [parameter()][String][AllowEmptyString()]$ObjectGuid,
        [parameter()][String][AllowEmptyString()]$ForceCreateTask,
        [parameter(Mandatory)][Ref]$returnObject
    )
    
    $TaskName = $TaskName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl +"api/v1/automationtasks?search=$TaskName&container=$AutomationContainer")
        $responseRaw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false) 
        $response = $responseRaw | Where-Object -filter {$_.name -eq $TaskName}
    
        if([string]::IsNullOrEmpty($response.automationTaskGuid) -or $ForceCreateTask -eq $true) {
            #Create Task

            $body = @{
                name                = $TaskName;
                useTemplate         = $UseTemplate;
                powerShellScript    = $PowershellScript;
                automationContainer = $AutomationContainer;
                objectGuid          = $ObjectGuid;
                variables           = (ConvertFrom-Json-WithEmptyArray($Variables));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/automationtasks/powershell")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $taskGuid = $response.automationTaskGuid

            Write-Information "Powershell task '$TaskName' created$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        } else {
            #Get TaskGUID
            $taskGuid = $response.automationTaskGuid
            Write-Warning "Powershell task '$TaskName' already exists$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        }
    } catch {
        Write-Error "Powershell task '$TaskName', message: $_"
    }

    $returnObject.Value = $taskGuid
}

function Invoke-HelloIDDatasource {
    param(
        [parameter(Mandatory)][String]$DatasourceName,
        [parameter(Mandatory)][String]$DatasourceType,
        [parameter(Mandatory)][String][AllowEmptyString()]$DatasourceModel,
        [parameter()][String][AllowEmptyString()]$DatasourceStaticValue,
        [parameter()][String][AllowEmptyString()]$DatasourcePsScript,        
        [parameter()][String][AllowEmptyString()]$DatasourceInput,
        [parameter()][String][AllowEmptyString()]$AutomationTaskGuid,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $DatasourceName = $DatasourceName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    $datasourceTypeName = switch($DatasourceType) { 
        "1" { "Native data source"; break} 
        "2" { "Static data source"; break} 
        "3" { "Task data source"; break} 
        "4" { "Powershell data source"; break}
    }
    
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/datasource/named/$DatasourceName")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
      
        if([string]::IsNullOrEmpty($response.dataSourceGUID)) {
            #Create DataSource
            $body = @{
                name               = $DatasourceName;
                type               = $DatasourceType;
                model              = (ConvertFrom-Json-WithEmptyArray($DatasourceModel));
                automationTaskGUID = $AutomationTaskGuid;
                value              = (ConvertFrom-Json-WithEmptyArray($DatasourceStaticValue));
                script             = $DatasourcePsScript;
                input              = (ConvertFrom-Json-WithEmptyArray($DatasourceInput));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
      
            $uri = ($script:PortalBaseUrl +"api/v1/datasource")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
              
            $datasourceGuid = $response.dataSourceGUID
            Write-Information "$datasourceTypeName '$DatasourceName' created$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        } else {
            #Get DatasourceGUID
            $datasourceGuid = $response.dataSourceGUID
            Write-Warning "$datasourceTypeName '$DatasourceName' already exists$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        }
    } catch {
      Write-Error "$datasourceTypeName '$DatasourceName', message: $_"
    }

    $returnObject.Value = $datasourceGuid
}

function Invoke-HelloIDDynamicForm {
    param(
        [parameter(Mandatory)][String]$FormName,
        [parameter(Mandatory)][String]$FormSchema,
        [parameter(Mandatory)][Ref]$returnObject
    )
    
    $FormName = $FormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/forms/$FormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }
    
        if(([string]::IsNullOrEmpty($response.dynamicFormGUID)) -or ($response.isUpdated -eq $true)) {
            #Create Dynamic form
            $body = @{
                Name       = $FormName;
                FormSchema = (ConvertFrom-Json-WithEmptyArray($FormSchema));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/forms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
    
            $formGuid = $response.dynamicFormGUID
            Write-Information "Dynamic form '$formName' created$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        } else {
            $formGuid = $response.dynamicFormGUID
            Write-Warning "Dynamic form '$FormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        }
    } catch {
        Write-Error "Dynamic form '$FormName', message: $_"
    }

    $returnObject.Value = $formGuid
}


function Invoke-HelloIDDelegatedForm {
    param(
        [parameter(Mandatory)][String]$DelegatedFormName,
        [parameter(Mandatory)][String]$DynamicFormGuid,
        [parameter()][Array][AllowEmptyString()]$AccessGroups,
        [parameter()][String][AllowEmptyString()]$Categories,
        [parameter(Mandatory)][String]$UseFaIcon,
        [parameter()][String][AllowEmptyString()]$FaIcon,
        [parameter()][String][AllowEmptyString()]$task,
        [parameter(Mandatory)][Ref]$returnObject
    )
    $delegatedFormCreated = $false
    $DelegatedFormName = $DelegatedFormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$DelegatedFormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }
    
        if([string]::IsNullOrEmpty($response.delegatedFormGUID)) {
            #Create DelegatedForm
            $body = @{
                name            = $DelegatedFormName;
                dynamicFormGUID = $DynamicFormGuid;
                isEnabled       = "True";
                useFaIcon       = $UseFaIcon;
                faIcon          = $FaIcon;
                task            = ConvertFrom-Json -inputObject $task;
            }
            if(-not[String]::IsNullOrEmpty($AccessGroups)) { 
                $body += @{
                    accessGroups    = (ConvertFrom-Json-WithEmptyArray($AccessGroups));
                }
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
    
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Information "Delegated form '$DelegatedFormName' created$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
            $delegatedFormCreated = $true

            $bodyCategories = $Categories
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$delegatedFormGuid/categories")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $bodyCategories
            Write-Information "Delegated form '$DelegatedFormName' updated with categories"
        } else {
            #Get delegatedFormGUID
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Warning "Delegated form '$DelegatedFormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
        }
    } catch {
        Write-Error "Delegated form '$DelegatedFormName', message: $_"
    }

    $returnObject.value.guid = $delegatedFormGuid
    $returnObject.value.created = $delegatedFormCreated
}


<# Begin: HelloID Global Variables #>
foreach ($item in $globalHelloIDVariables) {
	Invoke-HelloIDGlobalVariable -Name $item.name -Value $item.value -Secret $item.secret 
}
<# End: HelloID Global Variables #>


<# Begin: HelloID Data sources #>
<# Begin: DataSource "Entra-ID-MFA-Reset-generate-table-wildcard" #>
$tmpPsScript = @'
# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

function Get-ErrorMessage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber    = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line                = $ErrorObject.InvocationInfo.Line
            VerboseErrorMessage = $ErrorObject.Exception.Message
            AuditErrorMessage   = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.VerboseErrorMessage = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.VerboseErrorMessage = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.VerboseErrorMessage | ConvertFrom-Json)
            # Make sure to inspect the error result object and add only the error message as a FriendlyMessage.
            $httpErrorObj.VerboseErrorMessage = $errorDetailsObject.error
            $httpErrorObj.AuditErrorMessage = $errorDetailsObject.error.message
            if ($null -eq $httpErrorObj.AuditErrorMessage) {
                $httpErrorObj.AuditErrorMessage = $errorDetailsObject.error
            }
        }
        catch {
            $httpErrorObj.AuditErrorMessage = $httpErrorObj.VerboseErrorMessage
        }
        Write-Output $httpErrorObj
    }
}

try {
    $searchValue = $datasource.searchUser
    $searchQuery = "*$searchValue*"
          
    Write-Verbose "Generating Microsoft Graph API Access Token.."
    $baseUri = "https://login.microsoftonline.com/"
    $authUri = $baseUri + "$EntraTenantId/oauth2/token"
    $body = @{
        grant_type    = "client_credentials"
        client_id     = "$EntraAppId"
        client_secret = "$EntraAppSecret"
        resource      = "https://graph.microsoft.com"
    }

    $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
    $accessToken = $Response.access_token;
    Write-Information "Searching for: $searchQuery"
    
    #Add the authorization header to the request
    $authorization = @{
        Authorization  = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept         = "application/json";
    }

    $baseSearchUri = "https://graph.microsoft.com/"
    $searchUri = $baseSearchUri + "v1.0/users" + '?$select=Id,UserPrincipalName,displayName,department,jobTitle,companyName' + '&$top=999'

    $entraIDUsersResponse = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    $entraIDUsers = $entraIDUsersResponse.value
    while (![string]::IsNullOrEmpty($entraIDUsersResponse.'@odata.nextLink')) {
        $entraIDUsersResponse = Invoke-RestMethod -Uri $entraIDUsersResponse.'@odata.nextLink' -Method Get -Headers $authorization -Verbose:$false
        $entraIDUsers += $entraIDUsersResponse.value
    }  

    $users = foreach ($entraIDUser in $entraIDUsers) {
        if ($entraIDUser.displayName -like $searchQuery -or $entraIDUser.userPrincipalName -like $searchQuery) {
            $entraIDUser
        }
    }
    $users = $users | Sort-Object -Property DisplayName
    $resultCount = @($users).Count
    Write-Information "Result count: $resultCount"
        
    if ($resultCount -gt 0) {
        foreach ($user in $users) {
            $returnObject = @{
                Id                = $user.Id;
                UserPrincipalName = $user.UserPrincipalName;
                DisplayName       = $user.DisplayName;
                Department        = $user.Department;
                Title             = $user.JobTitle;
                Company           = $user.CompanyName
            }
            Write-Output $returnObject
        }
    }
}
catch {
    $ex = $PSItem

    $errorMessage = Get-ErrorMessage -ErrorObject $ex

    Write-Verbose "Error at Line [$($errorMessage.InvocationInfo.ScriptLineNumber)]: $($errorMessage.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))" 
    Write-Error "Error searching for Entra ID users. Error: $($errorMessage.AuditErrorMessage)"
}
'@ 
$tmpModel = @'
[{"key":"Department","type":0},{"key":"DisplayName","type":0},{"key":"Id","type":0},{"key":"UserPrincipalName","type":0},{"key":"Title","type":0},{"key":"Company","type":0}]
'@ 
$tmpInput = @'
[{"description":"","translateDescription":false,"inputFieldType":1,"key":"searchUser","type":0,"options":1}]
'@ 
$dataSourceGuid_0 = [PSCustomObject]@{} 
$dataSourceGuid_0_Name = @'
Entra-ID-MFA-Reset-generate-table-wildcard
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_0_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_0) 
<# End: DataSource "Entra-ID-MFA-Reset-generate-table-wildcard" #>
<# End: HelloID Data sources #>

<# Begin: Dynamic Form "Entra ID Account - MFA Reset" #>
$tmpSchema = @"
[{"label":"Select user account","fields":[{"key":"searchfield","templateOptions":{"label":"Search","placeholder":"Username or email address"},"type":"input","summaryVisibility":"Hide element","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"gridUsers","templateOptions":{"label":"Select user","required":true,"grid":{"columns":[{"headerName":"Display Name","field":"DisplayName"},{"headerName":"User Principal Name","field":"UserPrincipalName"},{"headerName":"Title","field":"Title"},{"headerName":"Department","field":"Department"},{"headerName":"Company","field":"Company"}],"height":300,"rowSelection":"single"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_0","input":{"propertyInputs":[{"propertyName":"searchUser","otherFieldValue":{"otherFieldKey":"searchfield"}}]}},"allowCsvDownload":true},"type":"grid","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":true}]}]
"@ 

$dynamicFormGuid = [PSCustomObject]@{} 
$dynamicFormName = @'
Entra ID Account - MFA Reset
'@ 
Invoke-HelloIDDynamicForm -FormName $dynamicFormName -FormSchema $tmpSchema  -returnObject ([Ref]$dynamicFormGuid) 
<# END: Dynamic Form #>

<# Begin: Delegated Form Access Groups and Categories #>
$delegatedFormAccessGroupGuids = @()
if(-not[String]::IsNullOrEmpty($delegatedFormAccessGroupNames)){
    foreach($group in $delegatedFormAccessGroupNames) {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/groups/$group")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
            $delegatedFormAccessGroupGuid = $response.groupGuid
            $delegatedFormAccessGroupGuids += $delegatedFormAccessGroupGuid
            
            Write-Information "HelloID (access)group '$group' successfully found$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormAccessGroupGuid })"
        } catch {
            Write-Error "HelloID (access)group '$group', message: $_"
        }
    }
    if($null -ne $delegatedFormAccessGroupGuids){
        $delegatedFormAccessGroupGuids = ($delegatedFormAccessGroupGuids | Select-Object -Unique | ConvertTo-Json -Depth 100 -Compress)
    }
}

$delegatedFormCategoryGuids = @()
foreach($category in $delegatedFormCategories) {
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories/$category")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        $response = $response | Where-Object {$_.name.en -eq $category}
        
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
        
        Write-Information "HelloID Delegated Form category '$category' successfully found$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    } catch {
        Write-Warning "HelloID Delegated Form category '$category' not found"
        $body = @{
            name = @{"en" = $category};
        }
        $body = ConvertTo-Json -InputObject $body -Depth 100

        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories")
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid

        Write-Information "HelloID Delegated Form category '$category' successfully created$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    }
}
$delegatedFormCategoryGuids = (ConvertTo-Json -InputObject $delegatedFormCategoryGuids -Depth 100 -Compress)
<# End: Delegated Form Access Groups and Categories #>

<# Begin: Delegated Form #>
$delegatedFormRef = [PSCustomObject]@{guid = $null; created = $null} 
$delegatedFormName = @'
Entra ID Account - MFA Reset
'@
$tmpTask = @'
{"name":"Entra ID Account - MFA Reset","script":"# Set TLS to accept TLS, TLS 1.1 and TLS 1.2\r\n[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12\r\n\r\n$VerbosePreference = \"SilentlyContinue\"\r\n$InformationPreference = \"Continue\"\r\n$WarningPreference = \"Continue\"\r\n\r\n# variables configured in form\r\n$userPrincipalName = $form.gridUsers.UserPrincipalName\r\n$id = $form.gridUsers.Id\r\n$displayName = $form.gridUsers.DisplayName\r\n\r\n# Debug\r\n# $id = '1b69b88a-5f6a-4010-84f2-ff4cabf5bd46'\r\n\r\n#region functions\r\nfunction Get-ErrorMessage {\r\n    [CmdletBinding()]\r\n    param (\r\n        [Parameter(Mandatory)]\r\n        [object]\r\n        $ErrorObject\r\n    )\r\n    process {\r\n        $httpErrorObj = [PSCustomObject]@{\r\n            ScriptLineNumber    = $ErrorObject.InvocationInfo.ScriptLineNumber\r\n            Line                = $ErrorObject.InvocationInfo.Line\r\n            VerboseErrorMessage = $ErrorObject.Exception.Message\r\n            AuditErrorMessage   = $ErrorObject.Exception.Message\r\n        }\r\n        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {\r\n            $httpErrorObj.VerboseErrorMessage = $ErrorObject.ErrorDetails.Message\r\n        }\r\n        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {\r\n            if ($null -ne $ErrorObject.Exception.Response) {\r\n                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()\r\n                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {\r\n                    $httpErrorObj.VerboseErrorMessage = $streamReaderResponse\r\n                }\r\n            }\r\n        }\r\n        try {\r\n            $errorDetailsObject = ($httpErrorObj.VerboseErrorMessage | ConvertFrom-Json)\r\n            # Make sure to inspect the error result object and add only the error message as a FriendlyMessage.\r\n            $httpErrorObj.VerboseErrorMessage = $errorDetailsObject.error\r\n            $httpErrorObj.AuditErrorMessage = $errorDetailsObject.error.message\r\n            if ($null -eq $httpErrorObj.AuditErrorMessage) {\r\n                $httpErrorObj.AuditErrorMessage = $errorDetailsObject.error\r\n            }\r\n        }\r\n        catch {\r\n            $httpErrorObj.AuditErrorMessage = $httpErrorObj.VerboseErrorMessage\r\n        }\r\n        Write-Output $httpErrorObj\r\n    }\r\n}\r\n\r\nfunction Remove-GraphAuthenticationMethod {\r\n    param (\r\n        [string]\r\n        $Type,\r\n\r\n        $Headers,\r\n\r\n        [string]\r\n        $UserId,\r\n\r\n        [string]\r\n        $MethodId,\r\n\r\n        [bool]\r\n        $Retry\r\n    )\r\n\r\n    process {\r\n        try {\r\n            $splatParams = @{\r\n                Method  = 'Delete'\r\n                Uri     = \"https://graph.microsoft.com/v1.0/users/$UserId/authentication/$Type/$MethodId\"\r\n                Headers = $Headers\r\n            }\r\n            $null = Invoke-RestMethod @splatParams -Verbose:$false\r\n            # Success is true, no retry\r\n            return $true\r\n        }\r\n        catch {\r\n            if (($_.ErrorDetails.Message -like \"*matches the user's current default authentication method*\") -and ($Retry -eq $false)) {\r\n                write-warning \"Couldn't revoke authentication method [$Type] [$MethodId] for Entra ID user [$id]. Retrying\"\r\n                # Success is false, retry\r\n                return $false\r\n            }\r\n            else {\r\n                Throw $_\r\n            }\r\n        }\r\n    }\r\n}\r\n#endregion functions\r\n\r\n\r\ntry {\r\n    #region Create authorization headers\r\n    $actionMessage = \"creating authorization headers\"\r\n    Write-Verbose \"Generating Microsoft Graph API Access Token..\"\r\n\r\n    $body = @{\r\n        grant_type    = \"client_credentials\"\r\n        client_id     = \"$EntraAppId\"\r\n        client_secret = \"$EntraAppSecret\"\r\n        resource      = \"https://graph.microsoft.com\"\r\n    }\r\n\r\n    $splatParamsToken = @{\r\n        Method      = 'POST'\r\n        Uri         = \"https://login.microsoftonline.com/$EntraTenantId/oauth2/token\"\r\n        Body        = $body\r\n        ContentType = 'application/x-www-form-urlencoded'\r\n    }\r\n    $Response = Invoke-RestMethod @splatParamsToken\r\n\r\n    $accessToken = $Response.access_token\r\n \r\n    #Add the authorization header to the request\r\n    $authorization = @{\r\n        Authorization  = \"Bearer $accesstoken\"\r\n        'Content-Type' = \"application/json\"\r\n        Accept         = \"application/json\"\r\n    }\r\n \r\n    #endregion Create authorization headers\r\n  \r\n    #region Get authentication methods\r\n    $actionMessage = \"getting authentication methods\"\r\n\r\n    Write-Verbose \"Getting authentication methods\"\r\n\r\n    $splatParamsGetAuthenticator = @{\r\n        Method  = 'Get'\r\n        Uri     = \"https://graph.microsoft.com/v1.0/users/$id/authentication/methods\"\r\n        Headers = $authorization\r\n    }\r\n    $responseGetAuthenticator = Invoke-RestMethod @splatParamsGetAuthenticator -Verbose:$false\r\n  \r\n    # Check if the response contains Microsoft Authenticator method\r\n    $microsoftAuthenticatorMethod = $responseGetAuthenticator.value | Where-Object { $_.'@odata.type' -eq \"#microsoft.graph.microsoftAuthenticatorAuthenticationMethod\" }\r\n\r\n    # Check if the response contains Phone Authentication method\r\n    $phoneAuthenticatorMethod = $responseGetAuthenticator.value | Where-Object { $_.'@odata.type' -eq \"#microsoft.graph.phoneAuthenticationMethod\" }\r\n\r\n    Write-Information \"Authentication methods successfully queried for Entra ID user [$userPrincipalName] [$id] successfully\"\r\n    #endregion Get authentication methods\r\n\r\n    #region Delete Phone Authentication method\r\n    $actionMessage = \"removing Phone Authentication method\"\r\n\r\n    if ($phoneAuthenticatorMethod) {\r\n        Write-Verbose \"Deleting current Phone Authentication method [$($phoneAuthenticatorMethod.phoneType)] with value [$($phoneAuthenticatorMethod.phoneNumber)] for account with id [$($id)]\"\r\n\r\n        $splatParamsDelMicrosoftAuthenticator = @{\r\n            Type     = 'phoneMethods'\r\n            Headers  = $authorization\r\n            UserId   = $id\r\n            MethodId = $phoneAuthenticatorMethod.id   \r\n            Retry    = $false            \r\n        }\r\n\r\n        $phoneAuthenticatorMethodSuccess = Remove-GraphAuthenticationMethod @splatParamsDelMicrosoftAuthenticator\r\n\r\n        if ($phoneAuthenticatorMethodSuccess) {\r\n            Write-Information \"Deleting current Phone Authentication method for Entra ID user [$userPrincipalName] [$id] successfully\"\r\n\r\n            $Log = @{\r\n                Action            = \"DeleteResource\" # optional. ENUM (undefined = default) \r\n                System            = \"Entra ID\" # optional (free format text) \r\n                Message           = \"Deleting current Phone Authentication method for Entra ID user [$userPrincipalName] [$id] successfully\" # required (free format text) \r\n                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                TargetDisplayName = $displayName # optional (free format text) \r\n                TargetIdentifier  = $([string]$id) # optional (free format text) \r\n            }\r\n            #send result back  \r\n            Write-Information -Tags \"Audit\" -MessageData $log\r\n        }\r\n    }\r\n    else {\r\n        Write-Verbose \"No Microsoft Authenticator method found for user [$id] [$userPrincipalName]\"       \r\n    }\r\n\r\n    #endregion Delete Phone Authentication method\r\n\r\n    #region Delete Microsoft Authenticator method\r\n    $actionMessage = \"removing Microsoft Authenticator method\"\r\n\r\n    if ($microsoftAuthenticatorMethod) {\r\n        Write-Verbose \"Deleting current Microsoft Authenticator method for account with id [$($id)]\"\r\n\r\n        $splatParamsDelMicrosoftAuthenticator = @{\r\n            Type     = 'microsoftAuthenticatorMethods'\r\n            Headers  = $authorization\r\n            UserId   = $id\r\n            MethodId = $microsoftAuthenticatorMethod.id \r\n            Retry    = $false             \r\n        }\r\n\r\n        $microsoftAuthenticatorMethodSuccess = Remove-GraphAuthenticationMethod @splatParamsDelMicrosoftAuthenticator\r\n\r\n        if ($microsoftAuthenticatorMethodSuccess) {\r\n            Write-Information \"Deleting current Microsoft Authenticator method for Entra ID user [$userPrincipalName] [$id] successfully\"\r\n\r\n            $Log = @{\r\n                Action            = \"DeleteResource\" # optional. ENUM (undefined = default) \r\n                System            = \"Entra ID\" # optional (free format text) \r\n                Message           = \"Deleting current Microsoft Authenticator method for Entra ID user [$userPrincipalName] [$id] successfully\" # required (free format text) \r\n                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                TargetDisplayName = $displayName # optional (free format text) \r\n                TargetIdentifier  = $([string]$id) # optional (free format text) \r\n            }\r\n            #send result back  \r\n            Write-Information -Tags \"Audit\" -MessageData $log\r\n        }\r\n    }\r\n    else {\r\n        Write-Verbose \"No Microsoft Authenticator method found for user [$id] [$userPrincipalName]\"\r\n    }\r\n\r\n    #endregion Delete Microsoft Authenticator method\r\n\r\n    #region Delete Phone Authentication method retry\r\n    $actionMessage = \"removing Phone Authentication method retry\"\r\n\r\n    if ($phoneAuthenticatorMethodSuccess -eq $false) {\r\n        Write-Verbose \"Retry deleting current Phone Authentication method [$($phoneAuthenticatorMethod.phoneType)] with value [$($phoneAuthenticatorMethod.phoneNumber)] for account with id [$($id)]\"\r\n\r\n        $splatParamsDelMicrosoftAuthenticator = @{\r\n            Type     = 'phoneMethods'\r\n            Headers  = $authorization\r\n            UserId   = $id\r\n            MethodId = $phoneAuthenticatorMethod.id\r\n            Retry    = $true            \r\n        }\r\n\r\n        $phoneAuthenticatorMethodSuccess = Remove-GraphAuthenticationMethod @splatParamsDelMicrosoftAuthenticator\r\n\r\n        Write-Information \"Retry deleting current Phone Authentication method for Entra ID user [$userPrincipalName] [$id] successfully\"\r\n\r\n        $Log = @{\r\n            Action            = \"DeleteResource\" # optional. ENUM (undefined = default) \r\n            System            = \"Entra ID\" # optional (free format text) \r\n            Message           = \"Retry deleting current Phone Authentication method for Entra ID user [$userPrincipalName] [$id] successfully\" # required (free format text) \r\n            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n            TargetDisplayName = $displayName # optional (free format text) \r\n            TargetIdentifier  = $([string]$id) # optional (free format text) \r\n        }\r\n        #send result back  \r\n        Write-Information -Tags \"Audit\" -MessageData $log\r\n    }\r\n\r\n    #endregion Delete Phone Authentication method retry\r\n\r\n    #region Delete Microsoft Authenticator method retry\r\n    $actionMessage = \"removing Microsoft Authenticator method retry\"\r\n\r\n    if ($microsoftAuthenticatorMethodSuccess -eq $false) {\r\n        Write-Verbose \"Retry deleting current Microsoft Authenticator method for account with id [$($id)]\"\r\n\r\n        $splatParamsDelMicrosoftAuthenticator = @{\r\n            Type     = 'microsoftAuthenticatorMethods'\r\n            Headers  = $authorization\r\n            UserId   = $id\r\n            MethodId = $microsoftAuthenticatorMethod.id    \r\n            Retry    = $true          \r\n        }\r\n\r\n        $microsoftAuthenticatorMethodSuccess = Remove-GraphAuthenticationMethod @splatParamsDelMicrosoftAuthenticator\r\n\r\n        Write-Information \"Retry deleting current Microsoft Authenticator method for Entra ID user [$userPrincipalName] [$id] successfully\"\r\n\r\n        $Log = @{\r\n            Action            = \"DeleteResource\" # optional. ENUM (undefined = default) \r\n            System            = \"Entra ID\" # optional (free format text) \r\n            Message           = \"Retry deleting current Microsoft Authenticator method for Entra ID user [$userPrincipalName] [$id] successfully\" # required (free format text) \r\n            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n            TargetDisplayName = $displayName # optional (free format text) \r\n            TargetIdentifier  = $([string]$id) # optional (free format text) \r\n        }\r\n        #send result back  \r\n        Write-Information -Tags \"Audit\" -MessageData $log\r\n    }\r\n\r\n    #endregion Delete Microsoft Authenticator method retry\r\n\r\n    #region no results found end of script\r\n    $actionMessage = \"no results found end of script\"\r\n    \r\n    if (($microsoftAuthenticatorMethod -eq $null) -and ($phoneAuthenticatorMethod -eq $null)) {\r\n        Write-Information \"No Microsoft Authenticator method and Phone Authentication method found for Entra ID user [$userPrincipalName] [$id]\"\r\n\r\n        $Log = @{\r\n            Action            = \"DeleteResource\" # optional. ENUM (undefined = default) \r\n            System            = \"Entra ID\" # optional (free format text) \r\n            Message           = \"No Microsoft Authenticator method and Phone Authentication method found for Entra ID user [$userPrincipalName] [$id]\" # required (free format text) \r\n            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n            TargetDisplayName = $displayName # optional (free format text) \r\n            TargetIdentifier  = $([string]$id) # optional (free format text) \r\n        }\r\n        #send result back  \r\n        Write-Information -Tags \"Audit\" -MessageData $log\r\n    }\r\n\r\n    #endregion no results found end of script\r\n}\r\ncatch {\r\n    $ex = $PSItem\r\n    $errorMessage = Get-ErrorMessage -ErrorObject $ex\r\n\r\n    Write-Verbose \"Error at Line [$($errorMessage.InvocationInfo.ScriptLineNumber)]: $($errorMessage.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))\" \r\n    Write-Error \"Error $actionMessage for Entra ID user [[$userPrincipalName] [$id]. Error: $($errorMessage.AuditErrorMessage)\"\r\n\r\n    $Log = @{\r\n        Action            = \"DeleteResource\" # optional. ENUM (undefined = default) \r\n        System            = \"Entra ID\" # optional (free format text) \r\n        Message           = \"Error $actionMessage for Entra ID user [$userPrincipalName] [$id]. Error: $($errorMessage.AuditErrorMessage)\" # required (free format text) \r\n        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $displayName # optional (free format text) \r\n        TargetIdentifier  = $([string]$id) # optional (free format text) \r\n    }\r\n    #send result back  \r\n    Write-Information -Tags \"Audit\" -MessageData $log\r\n}","runInCloud":true}
'@ 

Invoke-HelloIDDelegatedForm -DelegatedFormName $delegatedFormName -DynamicFormGuid $dynamicFormGuid -AccessGroups $delegatedFormAccessGroupGuids -Categories $delegatedFormCategoryGuids -UseFaIcon "True" -FaIcon "fa fa-key" -task $tmpTask -returnObject ([Ref]$delegatedFormRef) 
<# End: Delegated Form #>

