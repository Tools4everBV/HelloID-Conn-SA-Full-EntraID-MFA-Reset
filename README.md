# HelloID-Conn-SA-Full-Entra-ID-MFA-Reset

> [!IMPORTANT]
> This repository contains the connector and configuration code only. The implementer is responsible for acquiring the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements.

<p align="center">
  <img src="https://github.com/Tools4everBV/HelloID-Conn-SA-Full-Entra-ID-MFA-Reset/blob/main/Logo.png?raw=true">
</p>

## Table of contents

- [HelloID-Conn-SA-Full-Entra-ID-MFA-Reset](#helloid-conn-sa-full-entra-id-mfa-reset)
  - [Table of contents](#table-of-contents)
  - [Requirements](#requirements)
  - [Remarks](#remarks)
  - [Introduction](#introduction)
      - [Description](#description)
      - [Endpoints](#endpoints)
      - [Form Options](#form-options)
      - [Task Actions](#task-actions)
  - [Connector Setup](#connector-setup)
    - [Variable Library - User Defined Variables](#variable-library---user-defined-variables)
  - [Getting help](#getting-help)
  - [HelloID docs](#helloid-docs)

## Requirements
1. **HelloID Environment**:
   - Set up your _HelloID_ environment.
2. **Entra ID**:
   - App registration with `API permissions` of the type `Application`:
      -  `User.ReadWrite.All`
      - `UserAuthenticationMethod.ReadWrite.All`
   - The following information for the app registration is needed in HelloID:
      - `Application (client) ID`
      - `Directory (tenant) ID`
      - `Secret Value`

## Remarks
- The following methods are supported in this template `microsoftAuthenticatorAuthenticationMethod` and `phoneAuthenticationMethod`. Other methods can be added by enriching the action script.
- The default method should be removed last. But which method is default isn't reported by the graph API. For this reason, we retry removing a method one time. When retrying the method should be the last authentication method of the user and it will also be removed. When this also fails, an error is reported.

> [!IMPORTANT]
> If your organization uses other methods then `microsoftAuthenticatorAuthenticationMethod` and `phoneAuthenticationMethod` you should add them. If not the task can't delete the default method `microsoftAuthenticatorAuthenticationMethod` or `phoneAuthenticationMethod`

## Introduction

#### Description
_HelloID-Conn-SA-Full-Entra-ID-MFA-Reset_ is a template designed for use with HelloID Service Automation (SA) Delegated Forms. It can be imported into HelloID and customized according to your requirements. 

By using this delegated form, you can reset all MFA methods of an EntraID user. The following options are available:
 1. Search and select the Entra ID user
 2. The task will remove all the configured authentication methods

#### Endpoints
Entra Id provides a set of REST APIs that allow you to programmatically interact with its data. The API endpoints listed in the table below are used.

| Endpoint | Description                        |
| -------- | ---------------------------------- |
| users    | The user endpoint of the Graph API |

#### Form Options
The following options are available in the form:

1. **Lookup user**:
   - This Powershell data source runs an Entra ID Graph API query to search for matching Entra ID accounts.

#### Task Actions
The following actions will be performed based on user selections:

1. **Update UPN and Email in Active Directory**:
   - The current authentication methods of the selected user are retrieved and are stored in `$phoneAuthenticatorMethod` and `$microsoftAuthenticatorMethod`
   - If `$phoneAuthenticatorMethod` contains a value the `phoneMethods` will be removed. If it fails `$phoneAuthenticatorMethodSuccess` will be `$false`
   - If `$microsoftAuthenticatorMethod` contains a value the `microsoftAuthenticatorMethods` will be removed. If it fails `$microsoftAuthenticatorMethodSuccess` will be `$false`
   - If `$phoneAuthenticatorMethodSuccess` is `$false` the `phoneMethods` will be removed again. If it fails an error will be thrown
   - If `$microsoftAuthenticatorMethodSuccess` is `$false` the `microsoftAuthenticatorMethods` will be removed again. If it fails an error will be thrown

## Connector Setup
### Variable Library - User Defined Variables
The following user-defined variables are used by the connector. Ensure that you check and set the correct values required to connect to the API.

| Setting          | Description                                                     |
| ---------------- | --------------------------------------------------------------- |
| `EntraTenantId`  | The ID to the Tenant in Microsoft Entra ID                      |
| `EntraAppId`     | The ID to the App Registration in Microsoft Entra ID            |
| `EntraAppSecret` | The Client Secret to the App Registration in Microsoft Entra ID |

## Getting help
> [!TIP]
> _For more information on Delegated Forms, please refer to our [documentation](https://docs.helloid.com/en/service-automation/delegated-forms.html) pages_.

> [!TIP]
>  _If you need help, feel free to ask questions on our [forum](https://forum.helloid.com)_.

## HelloID docs
The official HelloID documentation can be found at: https://docs.helloid.com/