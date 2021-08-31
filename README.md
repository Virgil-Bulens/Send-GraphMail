# Send-GraphMail
### Azure Runbook that sends a simple mail with subject and content to a single recipient.

## REQUIREMENTS

- An App Registration with the Microsoft Graph API Mail.Send permission. The mailboxes this app has this permission on, should be restricted using the Exchange Online "New-ApplicationAccessPolicy" cmdlet.

- A licensed Exchange Online user or shared mailbox

- An active client secret for the App Registration

- A Key Vault containing the client secret

- A system assigned managed identity for the Automation Account this runbook is in. This identity should have at least the "Key Vault Secrets User" role on the client secret stored in the Key Vault, and the "Key Vault Reader" role on the Key Vault itself.

- Following variables stored in the Automation Account:
    - KeyVaultName
    - SecretName
    - TransactionalMailAppId
    - From
    - ExpirationWarningAddress

- Az.Accounts and Az.KeyVault modules imported in the Automation Account

---

## NOTES
- Author: Virgil Bulens
- Last Updated: 08/31/2021
- Version 1.0