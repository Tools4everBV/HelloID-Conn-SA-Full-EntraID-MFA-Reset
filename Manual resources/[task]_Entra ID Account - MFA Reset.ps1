# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form
$userPrincipalName = $form.gridUsers.UserPrincipalName
$id = $form.gridUsers.Id
$displayName = $form.gridUsers.DisplayName

# Debug
# $id = '1b69b88a-5f6a-4010-84f2-ff4cabf5bd46'

#region functions
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

function Remove-GraphAuthenticationMethod {
    param (
        [string]
        $Type,

        $Headers,

        [string]
        $UserId,

        [string]
        $MethodId,

        [bool]
        $Retry
    )

    process {
        try {
            $splatParams = @{
                Method  = 'Delete'
                Uri     = "https://graph.microsoft.com/v1.0/users/$UserId/authentication/$Type/$MethodId"
                Headers = $Headers
            }
            $null = Invoke-RestMethod @splatParams -Verbose:$false
            # Success is true, no retry
            return $true
        }
        catch {
            if (($_.ErrorDetails.Message -like "*matches the user's current default authentication method*") -and ($Retry -eq $false)) {
                write-warning "Couldn't revoke authentication method [$Type] [$MethodId] for Entra ID user [$id]. Retrying"
                # Success is false, retry
                return $false
            }
            else {
                Throw $_
            }
        }
    }
}
#endregion functions


try {
    #region Create authorization headers
    $actionMessage = "creating authorization headers"
    Write-Verbose "Generating Microsoft Graph API Access Token.."

    $body = @{
        grant_type    = "client_credentials"
        client_id     = "$EntraAppId"
        client_secret = "$EntraAppSecret"
        resource      = "https://graph.microsoft.com"
    }

    $splatParamsToken = @{
        Method      = 'POST'
        Uri         = "https://login.microsoftonline.com/$EntraTenantId/oauth2/token"
        Body        = $body
        ContentType = 'application/x-www-form-urlencoded'
    }
    $Response = Invoke-RestMethod @splatParamsToken

    $accessToken = $Response.access_token
 
    #Add the authorization header to the request
    $authorization = @{
        Authorization  = "Bearer $accesstoken"
        'Content-Type' = "application/json"
        Accept         = "application/json"
    }
 
    #endregion Create authorization headers
  
    #region Get authentication methods
    $actionMessage = "getting authentication methods"

    Write-Verbose "Getting authentication methods"

    $splatParamsGetAuthenticator = @{
        Method  = 'Get'
        Uri     = "https://graph.microsoft.com/v1.0/users/$id/authentication/methods"
        Headers = $authorization
    }
    $responseGetAuthenticator = Invoke-RestMethod @splatParamsGetAuthenticator -Verbose:$false
  
    # Check if the response contains Microsoft Authenticator method
    $microsoftAuthenticatorMethod = $responseGetAuthenticator.value | Where-Object { $_.'@odata.type' -eq "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" }

    # Check if the response contains Phone Authentication method
    $phoneAuthenticatorMethod = $responseGetAuthenticator.value | Where-Object { $_.'@odata.type' -eq "#microsoft.graph.phoneAuthenticationMethod" }

    Write-Information "Authentication methods successfully queried for Entra ID user [$userPrincipalName] [$id] successfully"
    #endregion Get authentication methods

    #region Delete Phone Authentication method
    $actionMessage = "removing Phone Authentication method"

    if ($phoneAuthenticatorMethod) {
        Write-Verbose "Deleting current Phone Authentication method [$($phoneAuthenticatorMethod.phoneType)] with value [$($phoneAuthenticatorMethod.phoneNumber)] for account with id [$($id)]"

        $splatParamsDelMicrosoftAuthenticator = @{
            Type     = 'phoneMethods'
            Headers  = $authorization
            UserId   = $id
            MethodId = $phoneAuthenticatorMethod.id   
            Retry    = $false            
        }

        $phoneAuthenticatorMethodSuccess = Remove-GraphAuthenticationMethod @splatParamsDelMicrosoftAuthenticator

        if ($phoneAuthenticatorMethodSuccess) {
            Write-Information "Deleting current Phone Authentication method for Entra ID user [$userPrincipalName] [$id] successfully"

            $Log = @{
                Action            = "DeleteResource" # optional. ENUM (undefined = default) 
                System            = "Entra ID" # optional (free format text) 
                Message           = "Deleting current Phone Authentication method for Entra ID user [$userPrincipalName] [$id] successfully" # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $displayName # optional (free format text) 
                TargetIdentifier  = $([string]$id) # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log
        }
    }
    else {
        Write-Verbose "No Microsoft Authenticator method found for user [$id] [$userPrincipalName]"       
    }

    #endregion Delete Phone Authentication method

    #region Delete Microsoft Authenticator method
    $actionMessage = "removing Microsoft Authenticator method"

    if ($microsoftAuthenticatorMethod) {
        Write-Verbose "Deleting current Microsoft Authenticator method for account with id [$($id)]"

        $splatParamsDelMicrosoftAuthenticator = @{
            Type     = 'microsoftAuthenticatorMethods'
            Headers  = $authorization
            UserId   = $id
            MethodId = $microsoftAuthenticatorMethod.id 
            Retry    = $false             
        }

        $microsoftAuthenticatorMethodSuccess = Remove-GraphAuthenticationMethod @splatParamsDelMicrosoftAuthenticator

        if ($microsoftAuthenticatorMethodSuccess) {
            Write-Information "Deleting current Microsoft Authenticator method for Entra ID user [$userPrincipalName] [$id] successfully"

            $Log = @{
                Action            = "DeleteResource" # optional. ENUM (undefined = default) 
                System            = "Entra ID" # optional (free format text) 
                Message           = "Deleting current Microsoft Authenticator method for Entra ID user [$userPrincipalName] [$id] successfully" # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $displayName # optional (free format text) 
                TargetIdentifier  = $([string]$id) # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log
        }
    }
    else {
        Write-Verbose "No Microsoft Authenticator method found for user [$id] [$userPrincipalName]"
    }

    #endregion Delete Microsoft Authenticator method

    #region Delete Phone Authentication method retry
    $actionMessage = "removing Phone Authentication method retry"

    if ($phoneAuthenticatorMethodSuccess -eq $false) {
        Write-Verbose "Retry deleting current Phone Authentication method [$($phoneAuthenticatorMethod.phoneType)] with value [$($phoneAuthenticatorMethod.phoneNumber)] for account with id [$($id)]"

        $splatParamsDelMicrosoftAuthenticator = @{
            Type     = 'phoneMethods'
            Headers  = $authorization
            UserId   = $id
            MethodId = $phoneAuthenticatorMethod.id
            Retry    = $true            
        }

        $phoneAuthenticatorMethodSuccess = Remove-GraphAuthenticationMethod @splatParamsDelMicrosoftAuthenticator

        Write-Information "Retry deleting current Phone Authentication method for Entra ID user [$userPrincipalName] [$id] successfully"

        $Log = @{
            Action            = "DeleteResource" # optional. ENUM (undefined = default) 
            System            = "Entra ID" # optional (free format text) 
            Message           = "Retry deleting current Phone Authentication method for Entra ID user [$userPrincipalName] [$id] successfully" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $displayName # optional (free format text) 
            TargetIdentifier  = $([string]$id) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }

    #endregion Delete Phone Authentication method retry

    #region Delete Microsoft Authenticator method retry
    $actionMessage = "removing Microsoft Authenticator method retry"

    if ($microsoftAuthenticatorMethodSuccess -eq $false) {
        Write-Verbose "Retry deleting current Microsoft Authenticator method for account with id [$($id)]"

        $splatParamsDelMicrosoftAuthenticator = @{
            Type     = 'microsoftAuthenticatorMethods'
            Headers  = $authorization
            UserId   = $id
            MethodId = $microsoftAuthenticatorMethod.id    
            Retry    = $true          
        }

        $microsoftAuthenticatorMethodSuccess = Remove-GraphAuthenticationMethod @splatParamsDelMicrosoftAuthenticator

        Write-Information "Retry deleting current Microsoft Authenticator method for Entra ID user [$userPrincipalName] [$id] successfully"

        $Log = @{
            Action            = "DeleteResource" # optional. ENUM (undefined = default) 
            System            = "Entra ID" # optional (free format text) 
            Message           = "Retry deleting current Microsoft Authenticator method for Entra ID user [$userPrincipalName] [$id] successfully" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $displayName # optional (free format text) 
            TargetIdentifier  = $([string]$id) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }

    #endregion Delete Microsoft Authenticator method retry

    #region no results found end of script
    $actionMessage = "no results found end of script"
    
    if (($microsoftAuthenticatorMethod -eq $null) -and ($phoneAuthenticatorMethod -eq $null)) {
        Write-Information "No Microsoft Authenticator method and Phone Authentication method found for Entra ID user [$userPrincipalName] [$id]"

        $Log = @{
            Action            = "DeleteResource" # optional. ENUM (undefined = default) 
            System            = "Entra ID" # optional (free format text) 
            Message           = "No Microsoft Authenticator method and Phone Authentication method found for Entra ID user [$userPrincipalName] [$id]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $displayName # optional (free format text) 
            TargetIdentifier  = $([string]$id) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }

    #endregion no results found end of script
}
catch {
    $ex = $PSItem
    $errorMessage = Get-ErrorMessage -ErrorObject $ex

    Write-Verbose "Error at Line [$($errorMessage.InvocationInfo.ScriptLineNumber)]: $($errorMessage.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))" 
    Write-Error "Error $actionMessage for Entra ID user [[$userPrincipalName] [$id]. Error: $($errorMessage.AuditErrorMessage)"

    $Log = @{
        Action            = "DeleteResource" # optional. ENUM (undefined = default) 
        System            = "Entra ID" # optional (free format text) 
        Message           = "Error $actionMessage for Entra ID user [$userPrincipalName] [$id]. Error: $($errorMessage.AuditErrorMessage)" # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $displayName # optional (free format text) 
        TargetIdentifier  = $([string]$id) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
