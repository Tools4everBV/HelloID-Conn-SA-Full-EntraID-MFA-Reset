# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

# Global variables
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# Fixed values
$retryCount = 0
$retryCountMax = 5

# Authentication methods to process
# For more information please check https://learn.microsoft.com/en-us/graph/api/resources/authenticationmethods-overview?view=graph-rest-1.0
$authenticationMethodsConfig = @{
    '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' = @{
        Type   = 'microsoftAuthenticatorMethods'
        Method = 'Microsoft Authenticator'
    }
    '#microsoft.graph.phoneAuthenticationMethod'                   = @{
        Type   = 'phoneMethods'
        Method = 'Phone Authentication'
    }
    '#microsoft.graph.emailAuthenticationMethod'                   = @{
        Type   = 'emailMethods'
        Method = 'Email Authentication'
    }
}

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Variables configured in form
$userPrincipalName = $form.gridUsers.UserPrincipalName
$id = $form.gridUsers.Id
$displayName = $form.gridUsers.DisplayName

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)
            if ($errorDetailsObject.error_description) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description
            }
            elseif ($errorDetailsObject.error.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)"
            }
            elseif ($errorDetailsObject.error.details.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)"
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}

function Remove-GraphAuthenticationMethod {
    param (
        [ValidateNotNullOrEmpty()]
        [string]
        $Type,

        [System.Collections.IDictionary]
        $Headers,

        [ValidateNotNullOrEmpty()]
        [string]
        $UserId,

        [ValidateNotNullOrEmpty()]
        [string]
        $MethodId,

        [ValidateNotNullOrEmpty()]
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
            if (($_.ErrorDetails.Message -like "*matches the user's current default authentication method*")) {
                Write-Warning "Couldn't revoke authentication method [$Type] [$MethodId] for Entra ID user [$UserId]. Retrying"
                # Success is false, retry
                return $false
            }
            elseif (($_.ErrorDetails.Message -like "*without first deleting alternate mobile number*")) {
                Write-Warning "Couldn't revoke authentication method [$Type] [$MethodId] for Entra ID user [$UserId]. Retrying"
                # Success is false, retry
                return $false
            }
            else {
                Throw $_
            }
        }
    }
}

function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Certificate,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TenantId
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$($AppId)"
            'sub' = "$($AppId)"
            'aud' = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-MSEntraCertificate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificateBase64String,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificatePassword
    )
    try {
        $rawCertificate = [system.convert]::FromBase64String($CertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
#endregion functions


try {
    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"
    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword
    Write-Verbose "Converted base64 certificate string to certificate object"

    # Create access token
    $actionMessage = "creating access token"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId

    # Create headers
    $actionMessage = "creating headers"
    $headers = @{
        "Authorization"    = "Bearer $($entraToken)"
        "Accept"           = "application/json"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }

    # Get authentication methods
    $actionMessage = "getting authentication methods for user [$userPrincipalName] with id [$id]"

    $splatParamsGetAuthenticator = @{
        Method  = 'Get'
        Uri     = "https://graph.microsoft.com/v1.0/users/$id/authentication/methods"
        Headers = $headers
    }
    $responseGetAuthenticator = Invoke-RestMethod @splatParamsGetAuthenticator -Verbose:$false
    Write-Information "Queried authentication methods for Entra ID user [$userPrincipalName] with id [$id]. Total result count: $(($responseGetAuthenticator.value | Measure-Object).Count)"

    # Filter authenticators based on configured authentication methods
    $configuredTypes = ($authenticationMethodsConfig.Values | ForEach-Object { $_.Method }) -join ', '
    $actionMessage = "filtering authentication methods to configured types [$configuredTypes] for Entra ID user [$userPrincipalName] with id [$id]"
    $authenticators = $responseGetAuthenticator.value | Where-Object { $authenticationMethodsConfig.ContainsKey($_.'@odata.type') }
    Write-Information "Filtered authentication methods to configured types [$configuredTypes] for Entra ID user [$userPrincipalName] with id [$id]. Result count: $(($authenticators | Measure-Object).Count)"

    # Delete authentication methods with retry logic
    $actionMessage = "deleting authentication methods for user [$userPrincipalName] with id [$id]"
    $authenticators | Add-Member -MemberType NoteProperty -Name "retry" -Value $false -Force
    $authenticators | Add-Member -MemberType NoteProperty -Name "success" -Value $false -Force
    do {
        $doUntilSuccess = $true
        foreach ($authenticator in $authenticators) {
            $methodId = $authenticator.id

            # Get authentication method configuration
            $authConfig = $authenticationMethodsConfig[$authenticator.'@odata.type']
            if ($null -ne $authConfig) {
                $type = $authConfig.Type
                $method = $authConfig.Method
            }
            else {
                # If no configured method is found then skip (success = $true)
                $authenticator.success = $true
                continue
            }

            if (-not $authenticator.success) {
                $actionMessage = "deleting current $method method for user with id [$($id)]"
                $actionsNeeded = $true

                $splatParamsDelAuthenticator = @{
                    Type     = $type
                    Headers  = $headers
                    UserId   = $id
                    MethodId = $methodId 
                    Retry    = $authenticator.retry         
                }

                $authenticator.success = Remove-GraphAuthenticationMethod @splatParamsDelAuthenticator

                if ($authenticator.success) {
                    $Log = @{
                        Action            = "DeleteResource" # optional. ENUM (undefined = default) 
                        System            = "EntraID" # optional (free format text) 
                        Message           = "Deleted current $method method for Entra ID user [$userPrincipalName] with id [$id]" # required (free format text) 
                        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $displayName # optional (free format text) 
                        TargetIdentifier  = $([string]$id) # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log
                }
                else {
                    $authenticator.retry = $true
                    $retryCount++
                    $doUntilSuccess = $false
                }
            }
        }
    } until ($doUntilSuccess -or $retryCount -gt $retryCountMax)

    if ($actionsNeeded -ne $true) {
        $Log = @{
            Action            = "DeleteResource" # optional. ENUM (undefined = default) 
            System            = "EntraID" # optional (free format text) 
            Message           = "No authentication method found that needs to be deleted for Entra ID user [$userPrincipalName] with id [$id]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $displayName # optional (free format text) 
            TargetIdentifier  = $([string]$id) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }

    $Log = @{
        Action            = "DeleteResource" # optional. ENUM (undefined = default) 
        System            = "EntraID" # optional (free format text) 
        Message           = $auditMessage # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $displayName # optional (free format text) 
        TargetIdentifier  = $([string]$id) # optional (free format text) 
    }
    Write-Information -Tags "Audit" -MessageData $log
    Write-Warning $warningMessage
    Write-Error $auditMessage
}