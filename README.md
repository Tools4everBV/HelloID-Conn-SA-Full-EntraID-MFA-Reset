# HelloID-Conn-SA-Full-EntraID-MFA-Reset

| :information_source: Information |
| :------------------------------- |
| This repository contains the connector and configuration code only. The implementer is responsible for acquiring the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements. |

## Description
_HelloID-Conn-SA-Full-EntraID-MFA-Reset_ is a template designed for use with HelloID Service Automation (SA) Delegated Forms. It can be imported into HelloID and customized according to your requirements.

By using this delegated form, you can reset all MFA methods of a Microsoft Entra ID user. The following options are available:
 1. Search and select the user
 2. The task removes the configured authentication methods

## Getting started
### Requirements

#### App Registration & Certificate Setup

Before implementing this connector, make sure to configure a Microsoft Entra ID, an App Registration. During the setup process, you’ll create a new App Registration in the Entra portal, assign the necessary API permissions (such as user and group read/write), and generate and assign a certificate.

Follow the official Microsoft documentation for creating an App Registration and setting up certificate-based authentication:
- [App-only authentication with certificate (Exchange Online)](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#set-up-app-only-authentication)

#### HelloID-specific configuration

Once you have completed the Microsoft setup and followed their best practices, configure the following HelloID-specific requirements.

- **API Permissions** (Application permissions):
  - `User.ReadWrite.All`
  - `Group.ReadWrite.All`
  - `GroupMember.ReadWrite.All`
  - `UserAuthenticationMethod.ReadWrite.All`
  - `User.EnableDisableAccount.All`
  - `User-PasswordProfile.ReadWrite.All`
  - `User-Phone.ReadWrite.All`
- **Certificate:**
  - Upload the public key file (.cer) in Entra ID
  - Provide the certificate as a Base64 string in HelloID. For instructions on creating the certificate and obtaining the base64 string, refer to our forum post: [Setting up a certificate for Microsoft Graph API in HelloID connectors](https://forum.helloid.com/forum/helloid-provisioning/5338-instruction-setting-up-a-certificate-for-microsoft-graph-api-in-helloid-connectors#post5338)


### Connection settings

The following user-defined variables are used by the connector.

| Setting                       | Description                                                     | Mandatory |
| ----------------------------- | --------------------------------------------------------------- | --------- |
| EntraIdTenantId               | The Directory (tenant) ID in Microsoft Entra ID                 | Yes       |
| EntraIdAppId                  | The Application (client) ID in Microsoft Entra ID               | Yes       |
| EntraIdCertificateBase64String| Base64-encoded certificate used for client assertion            | Yes       |
| EntraIdCertificatePassword    | Password for the provided certificate (if applicable)           | Yes       |

## Remarks

### Supported Authentication Methods
- This template supports `microsoftAuthenticatorAuthenticationMethod` and `phoneAuthenticationMethod`. Other methods can be added by enriching the task script.

### Default Method Removal Retry
- The Graph API does not indicate which method is the default. The task retries removal once when the default method blocks deletion. On retry, the last remaining method is removed. If this also fails, an error is reported.

## Development resources

### API endpoints

The following endpoints are used by the connector

| Endpoint                                          | Description                                         |
| ------------------------------------------------- | --------------------------------------------------- |
| /users                                            | Retrieve user information                           |
| /users/{id}/authentication/methods                | List a user's authentication methods                |
| /users/{id}/authentication/phoneMethods/{methodId}| Remove a phone authentication method                |
| /users/{id}/authentication/microsoftAuthenticatorMethods/{methodId} | Remove a Microsoft Authenticator method |

### API documentation

- Microsoft Graph: Authentication methods overview: https://learn.microsoft.com/graph/api/resources/authenticationmethods-overview
- Microsoft Graph: Users API: https://learn.microsoft.com/graph/api/resources/users

## Getting help
> :bulb: **Tip:**  
> _For more information on Delegated Forms, please refer to our [documentation](https://docs.helloid.com/en/service-automation/delegated-forms.html) pages_.

## HelloID docs
The official HelloID documentation can be found at: https://docs.helloid.com/
