# Microsoft 365 Daily Activity

A **Blazor Server** web app built on **.NET 10** that fetches and displays your Microsoft 365 activity for any selected date — Teams messages, calendar events, emails, and Azure DevOps commits — all in one place via the Microsoft Graph API.

---

## Features

| Source | What is fetched |
|---|---|
| 💬 **Teams Messages** | 1:1 chats, group chats, and channel messages |
| 📅 **Calendar** | All meetings and events for the selected day |
| 📧 **Email** | Sent and received messages from Outlook |
| 🔧 **DevOps Commits** | Your commits across all Azure DevOps projects and repos |

- Interactive date picker — browse any past date
- Summary cards showing sent/received counts at a glance
- CSV export saved to your Desktop
- Microsoft sign-in via interactive browser (OAuth 2.0, token cached in memory)

---

## Tech Stack

| Layer | Technology |
|---|---|
| Framework | .NET 10 / ASP.NET Core |
| UI | Blazor Server + Bootstrap 5 |
| Microsoft 365 API | Microsoft Graph SDK v5 |
| Authentication | Azure.Identity — `InteractiveBrowserCredential` |
| DevOps API | Azure DevOps REST API v7.1 |

---

## Project Structure

```
TeamsMessageFetcher/
├── Program.cs                      # App startup & DI registration
├── App.razor                       # HTML shell (Bootstrap, favicon)
├── Routes.razor                    # Blazor router
├── _Imports.razor                  # Global @using directives
├── appsettings.json                # TenantId, ClientId, DevOps config
│
├── Configuration/
│   └── AppConfig.cs                # Reads config from appsettings / env vars
│
├── Models/
│   ├── TeamMessage.cs              # Teams chat/channel message
│   ├── MailMessage.cs              # Email (sent or received)
│   ├── CalendarEvent.cs            # Calendar event / meeting
│   └── DevOpsCommit.cs            # Azure DevOps commit
│
├── Graph/
│   ├── GraphClientFactory.cs       # Builds the authenticated GraphServiceClient
│   └── GraphPaginator.cs          # Generic Graph pagination helper
│
├── Fetchers/
│   ├── MessageFetcher.cs           # Fetches Teams chats + channel messages
│   ├── MailFetcher.cs              # Fetches sent and received email
│   ├── CalendarFetcher.cs          # Fetches calendar events
│   └── DevOpsFetcher.cs           # Fetches DevOps commits via REST API
│
├── Services/
│   ├── DailyActivityService.cs     # Orchestrates all fetchers; used by the UI
│   └── MessageExporter.cs         # Exports data to CSV
│
├── Components/
│   ├── Layout/
│   │   └── MainLayout.razor        # Top navbar + page wrapper
│   └── Pages/
│       └── Home.razor              # Main page (date picker, cards, sections)
│
└── wwwroot/
    ├── app.css                     # Custom styles
    └── favicon.svg                 # App icon
```

---

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- A Microsoft 365 account (Teams + Outlook + Calendar)
- An Azure AD App Registration — see setup below
- *(Optional)* An Azure DevOps organisation + Personal Access Token

---

## Step 1 — Azure AD App Registration

1. Go to [portal.azure.com](https://portal.azure.com)
2. Navigate to **Microsoft Entra ID** → **App registrations** → **New registration**
3. Fill in:
   - **Name:** `TeamsMessageFetcher` (or any name)
   - **Supported account types:** Single tenant
   - **Redirect URI:** leave blank for now
4. Click **Register**, then copy the **Application (client) ID** and **Directory (tenant) ID**

### Authentication settings

In your app registration → **Authentication**:

- Click **Add a platform** → choose **Mobile and desktop applications**
- Add `http://localhost` as a redirect URI → **Save**
- Under **Advanced settings**, enable **Allow public client flows** → **Save**

### API Permissions

Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**:

| Permission | Purpose |
|---|---|
| `User.Read` | Read your profile and get your user ID |
| `Chat.Read` | Read your 1:1 and group chat messages |
| `Chat.ReadBasic` | List your chats |
| `ChannelMessage.Read.All` | Read messages in Teams channels |
| `Team.ReadBasic.All` | List your joined teams and channels |
| `Calendars.Read` | Read your calendar events |
| `Mail.Read` | Read your received emails |
| `Mail.ReadSent` | Read your sent emails |

Click **Grant admin consent** (requires an admin, or request consent from yours).

> **Note:** `ChannelMessage.Read.All` typically requires admin consent in most organisations.

---

## Step 2 — Configure the App

Edit `appsettings.json` with your IDs:

```json
{
  "TenantId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "ClientId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",

  "DevOpsOrganization": "your-org-name",
  "DevOpsPat": "your-personal-access-token"
}
```

| Key | Required | Description |
|---|---|---|
| `TenantId` | ✅ | Your Azure AD Directory (tenant) ID |
| `ClientId` | ✅ | Your App Registration's Application (client) ID |
| `DevOpsOrganization` | ⬜ Optional | Your Azure DevOps organisation name (e.g. `contoso`) |
| `DevOpsPat` | ⬜ Optional | A PAT with **Code (read)** scope. Commits are skipped if absent. |

You can also use environment variables instead of `appsettings.json`:

```bash
# Windows
set TEAMS_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
set TEAMS_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

# macOS / Linux
export TEAMS_TENANT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
export TEAMS_CLIENT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

### Creating an Azure DevOps PAT

1. Go to `https://dev.azure.com/{your-org}` → top-right user icon → **Personal access tokens**
2. Click **New Token**
3. Set **Scopes** → **Code** → **Read**
4. Copy the generated token into `appsettings.json`

---

## Step 3 — Run

```bash
# Restore packages
dotnet restore

# Start the web server
dotnet run
```

Then open your browser at **`https://localhost:5001`** (or `http://localhost:5000`).

On first use, clicking **Fetch Activity** opens a Microsoft sign-in browser tab. After signing in, the token is cached in memory for the session — subsequent fetches won't prompt again.

---

## Usage

1. **Select a date** using the date picker (defaults to today)
2. Click **🔍 Fetch Activity** — a spinner shows while data is loading
3. View your activity grouped by source in the four sections below the summary cards
4. Click **💾 Export CSV** to save a `daily_summary_YYYY-MM-DD.csv` to your Desktop

---

## Notes & Limitations

| Topic | Detail |
|---|---|
| **Chat date filtering** | The Graph API does not support `$filter` on chat messages — they are fetched in descending order and filtered client-side. Very active chats may be slower to process. |
| **Token caching** | `InteractiveBrowserCredential` caches the token in memory. Restarting the app requires signing in again. |
| **Rate limits** | The Graph API applies per-user throttling. If you have many teams or channels, occasional `429` responses are automatically retried by the SDK. |
| **Admin consent** | `ChannelMessage.Read.All` and `Mail.Read` typically require admin consent. Contact your Microsoft 365 administrator if permissions are denied. |
| **DevOps commits** | Commits are matched by the email address on your Microsoft 365 account (`me.Mail`). Ensure your DevOps commit author email matches. |
