# HelloID-Conn-SA-Full-EntraID-MFA-Reset

| :information_source: Information |
| :------------------------------- |
| This repository contains the connector and configuration code only. The implementer is responsible for acquiring the connection details such as certificate, tenant ID, and application ID. You might need administrator consent and to configure an App Registration in Microsoft Entra ID before implementing this connector. Please contact the client's application owner to coordinate the requirements. |

## Description

HelloID-Conn-SA-Full-EntraID-MFA-Reset is a delegated form designed for use with HelloID Service Automation (SA). It can be imported into HelloID and customized according to your requirements.

By using this delegated form, you can reset all configured MFA authentication methods for a Microsoft Entra ID user. The following options are available:

1. Search for and select the target Microsoft Entra ID user account (wildcard search by display name, UserPrincipalName, or mail).
2. The task removes all configured authentication methods for the selected user account.

## Getting started
### Requirements

#### App Registration & Certificate Setup

Before implementing this connector, make sure to configure a Microsoft Entra ID App Registration. During the setup process, you'll create a new App Registration in the Entra portal, assign the necessary API permissions (such as user and authentication method read/write), and generate and assign a certificate.

Follow the official Microsoft documentation for creating an App Registration and setting up certificate-based authentication:

- [App-only authentication with certificate](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#set-up-app-only-authentication)

#### HelloID-specific configuration

Once you have completed the Microsoft setup and followed their best practices, configure the following HelloID-specific requirements.

- **API Permissions** (Application permissions):
  - `UserAuthenticationMethod.ReadWrite.All` - To manage user authentication methods
  - `User.Read.All` - To read user information
- **Certificate Base64 encoded string:**
  - Base64 encoded string of the certificate assigned to the app registration. For instructions on creating the certificate and obtaining the base64 string, refer to our forum post: [Setting up a certificate for Microsoft Graph API in HelloID connectors](https://forum.helloid.com/forum/helloid-provisioning/5338-instruction-setting-up-a-certificate-for-microsoft-graph-api-in-helloid-connectors#post5338)


### Connection settings

The following global variables must be configured in HelloID when importing and configuring the delegated form.

| Setting                        | Description                                                                | Mandatory |
| ------------------------------ | -------------------------------------------------------------------------- | --------- |
| EntraIdTenantId                | The unique identifier (ID) of the tenant in Microsoft Entra ID             | Yes       |
| EntraIdAppId                   | The unique identifier (ID) of the App Registration in Microsoft Entra ID   | Yes       |
| EntraIdCertificateBase64String | The Base64-encoded string representation of the app certificate            | Yes       |
| EntraIdCertificatePassword     | The password associated with the app certificate                           | Yes       |

## Remarks

### Supported Authentication Methods

The connector is configured to process the following authentication method types:
- **Microsoft Authenticator** (`microsoftAuthenticatorAuthenticationMethod`)
- **Phone Authentication** (`phoneAuthenticationMethod`)
- **Email Authentication** (`emailAuthenticationMethod`)

Additional authentication methods can be added by extending the `$authenticationMethodsConfig` hashtable at the top of the task script. For a complete list of supported authentication methods, refer to the [Microsoft Graph authentication methods overview](https://learn.microsoft.com/en-us/graph/api/resources/authenticationmethods-overview?view=graph-rest-1.0).

### Default Method Removal Retry

The Microsoft Graph API does not explicitly indicate which authentication method is set as the user's default. When attempting to remove the default method, the API returns an error indicating it cannot be removed. 

The connector implements an automatic retry mechanism:
1. On first attempt, all configured methods are removed
2. If a method fails due to being the default, it is marked for retry
3. The connector automatically retries removal of failed methods (up to 5 retry attempts)
4. On retry, the previously default method can typically be removed as another method has become the new default

This also handles the scenario where a phone number cannot be removed without first deleting an alternate mobile number.

### Certificate-Based Authentication

The connector uses certificate-based authentication to generate JSON Web Tokens (JWT) for secure communication with Microsoft Graph API. The certificate is converted from a base64 string and used to sign the JWT assertion for OAuth2 authentication, providing enhanced security compared to client secret authentication.

### Error Handling

- **Configurable Methods**: Only authentication methods defined in `$authenticationMethodsConfig` are processed. Other method types are automatically skipped.
- **Retry Logic**: Methods that fail due to being the default or other temporary issues are automatically retried with a configurable maximum retry count.
- **Comprehensive Logging**: All operations are logged with detailed audit messages indicating success or failure, including specific method types and user information.

## Development resources

### API endpoints

The following Microsoft Graph API endpoints are used by the connector:

| Endpoint                                                                  | Description                              |
| ------------------------------------------------------------------------- | ---------------------------------------- |
| /v1.0/users                                                               | List users                               |
| /v1.0/users/{id}/authentication/methods                                   | List user authentication methods         |
| /v1.0/users/{id}/authentication/phoneMethods/{methodId}                   | Remove phone authentication method       |
| /v1.0/users/{id}/authentication/microsoftAuthenticatorMethods/{methodId}  | Remove Microsoft Authenticator method    |

### API documentation

- [List users](https://learn.microsoft.com/en-us/graph/api/user-list)
- [List authentication methods](https://learn.microsoft.com/en-us/graph/api/authentication-list-methods)
- [Authentication methods overview](https://learn.microsoft.com/en-us/graph/api/resources/authenticationmethods-overview?view=graph-rest-1.0)
- [Delete phoneAuthenticationMethod](https://learn.microsoft.com/en-us/graph/api/phoneauthenticationmethod-delete)
- [Delete microsoftAuthenticatorAuthenticationMethod](https://learn.microsoft.com/en-us/graph/api/microsoftauthenticatorauthenticationmethod-delete)

## Getting help

> :bulb: **Tip:** For more information on Delegated Forms, please refer to our [documentation](https://docs.helloid.com/en/service-automation/delegated-forms.html) pages.

## HelloID docs

The official HelloID documentation can be found at: [https://docs.helloid.com/](https://docs.helloid.com/)
