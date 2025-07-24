
### 📘 OutlookCategoryCreator

This PowerShell script allows you to **bulk-create Outlook categories** using the Microsoft Graph API and delegated permissions.

It is tailored for **Microsoft 365 mailboxes** (e.g., `contact @ cloudmechanics.ae`) and is useful for email automation and mailbox organization in event management or sponsorship workflows.

---

### ✅ Features

* Creates multiple custom categories in Outlook.
* Uses Microsoft Graph API (`/me/outlook/masterCategories` endpoint).
* Handles token generation using `MSAL.PS` module.
* Customizable list of labels.
* Works with delegated user authentication.

---

### 🚀 Prerequisites

* PowerShell 7+
* Microsoft Graph application with delegated permissions:

  * `MailboxSettings.ReadWrite`
  * `User.Read`
* Registered App with:

  * Redirect URI: `http://localhost`
  * Public client (no client secret)
* Module: `MSAL.PS`
  Install with:

  ```powershell
  Install-Module MSAL.PS -Scope CurrentUser
  ```

---

### 🧾 Usage

1. Clone the repo or download the `.zip`.
2. Edit the variables:

   ```powershell
   $clientId = "<Your App Client ID>"
   $tenantId = "<Your Tenant ID>"
   $userEmail = "user@example.ae"
   ```
3. Run the script:

   ```powershell
   ./Create-OutlookCategories.ps1
   ```

---

### 🏷️ Default Categories

The script creates the following labels:

```plaintext
Action Required
Newsletter
Spam
Receipt
Social
Event Invite
Subscription Renewal
System Notification
Autoresponse
```

---

### 👤 Author

**Mahmoud A. Atallah**
Azure Solutions Lead | Microsoft MVP
* 📧 [mahmoud@cloudmechanics.ae](mailto:mahmoud@cloudmechanics.ae)
* 🌐 [cloudmechanics.ae](https://cloudmechanics.ae)
* 🔗 [linkedin.com/in/mahmoudatallah](https://linkedin.com/in/mahmoudatallah)

