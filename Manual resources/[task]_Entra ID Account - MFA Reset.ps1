# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form
$userPrincipalName = $form.gridUsers.UserPrincipalName
$id = $form.gridUsers.Id
$displayName = $form.gridUsers.DisplayName

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

function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Certificate
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
            'iss' = "$entraidappid"
            'sub' = "$entraidappid"
            'aud' = "https://login.microsoftonline.com/$EntraIdTenantId/oauth2/token"
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
	
	# Extract the private key from the certificate
        if (-not $Certificate.HasPrivateKey -or -not $Certificate.PrivateKey) {
            throw "The certificate does not have a private key."
        }

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $entraidappid
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$EntraIdTenantId/oauth2/token"
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
    param()
    try {
        $rawCertificate = [system.convert]::FromBase64String($EntraIdCertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $EntraIdCertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
#endregion functions


try {
    # Setup Connection with Entra/Exo
    Write-Verbose 'connecting to MS-Entra'
    $certificate = Get-MSEntraCertificate
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate
    
    #Add the authorization header to the request
    $authorization = @{
        Authorization = "Bearer $entraToken";
        'Content-Type' = "application/json";
        Accept = "application/json";
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
