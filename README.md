# Huginn Outlook Agent

A Huginn agent gem that enables integration with Microsoft Outlook email via the Microsoft Graph API. This agent provides both email receiving and sending capabilities.

## Features

- **OutlookAgent**: Monitor Outlook inbox for new emails and create Huginn events
- **OutlookSenderAgent**: Send emails through Outlook based on Huginn events
- Support for HTML and plain text emails
- Liquid templating support for dynamic email content
- OAuth 2.0 authentication with Microsoft Graph API
- Configurable email filtering and processing

## Installation

This gem is run as part of the [Huginn](https://github.com/huginn/huginn) project. If you haven't already, follow the [Getting Started](https://github.com/huginn/huginn#getting-started) instructions there.

Add this string to your Huginn's .env `ADDITIONAL_GEMS` configuration:

```ruby
huginn_outlook_agent
# when only using this agent gem it should look like this:
ADDITIONAL_GEMS=huginn_outlook_agent
```

And then execute:

```bash
$ bundle
```

**Note:** This agent uses `oauth2 ~> 1.4` to maintain compatibility with existing Huginn dependencies like `dropbox-api`. The OAuth functionality works identically with this version.

## Microsoft Graph API Setup

Before using this agent, you need to set up Microsoft Graph API access:

1. **Register an Application in Azure AD**
   - Go to [Azure Portal](https://portal.azure.com)
   - Navigate to "Azure Active Directory" → "App registrations"
   - Click "New registration"
   - Give it a name (e.g., "Huginn Outlook Integration")
   - Choose **"Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant)"**
   - Note the Application (client) ID and Directory (tenant) ID

2. **Configure API Permissions**
   - Go to "API permissions" → "Add a permission"
   - Select "Microsoft Graph"
   - Add these permissions:
     - `Mail.Read` (for receiving emails)
     - `Mail.Send` (for sending emails)
   - Grant admin consent

3. **Create Client Secret**
   - Go to "Certificates & secrets"
   - Click "New client secret"
   - Copy the secret value immediately (it won't be shown again)

4. **Configure Agent Authentication**
   
   The agents support two authentication methods:

   **Option A: Built-in OAuth (Recommended)**
   - Set `auth_method` to `"oauth"`
   - Provide `client_id`, `client_secret`, and `tenant_id`
   - The agent will automatically handle token acquisition and refresh
   - No external token management required

   **Option B: Manual Access Token**
   - Set `auth_method` to `"token"` (or leave blank)
   - Provide a manually obtained `access_token`
   - You'll need to refresh the token when it expires (~1 hour)

   **Important Notes:**
   - OAuth is the recommended method for production use
   - The agent handles token refresh automatically with OAuth
   - Store credentials securely using Huginn's credentials system

## Usage

### OutlookAgent (Receiving Emails)

Configure the agent with these options:

**Using OAuth (Recommended):**
```json
{
  "mode": "receive",
  "auth_method": "oauth",
  "client_id": "your_client_id",
  "client_secret": "your_client_secret",
  "tenant_id": "your_tenant_id",
  "folder": "inbox",
  "since": "2023-01-01T00:00:00Z",
  "mark_as_read": false
}
```

**Using Manual Access Token:**
```json
{
  "mode": "receive",
  "auth_method": "token",
  "access_token": "your_access_token_here",
  "folder": "inbox",
  "since": "2023-01-01T00:00:00Z",
  "mark_as_read": false
}
```

**Options:**
- `mode`: Must be "receive"
- `auth_method`: "oauth" (recommended) or "token"
- `client_id`: Azure AD application client ID (for OAuth)
- `client_secret`: Azure AD application client secret (for OAuth)
- `tenant_id`: Azure AD directory tenant ID (for OAuth)
- `access_token`: Microsoft Graph API access token (for token method)
- `folder`: One of "inbox", "sent", "drafts", "deleted"
- `since`: ISO 8601 datetime to filter emails (optional)
- `mark_as_read`: Whether to mark processed emails as read (default: false)

**Event Payload:**
```json
{
  "id": "email_id",
  "subject": "Email Subject",
  "from": "sender@example.com",
  "to": ["recipient@example.com"],
  "cc": ["cc@example.com"],
  "body": "Email content (HTML or text)",
  "received_at": "2023-01-01T12:00:00Z",
  "is_read": false
}
```

### OutlookSenderAgent (Sending Emails)

Configure the agent with these options:

**Using OAuth (Recommended):**
```json
{
  "auth_method": "oauth",
  "client_id": "your_client_id",
  "client_secret": "your_client_secret",
  "tenant_id": "your_tenant_id",
  "to": "recipient@example.com",
  "subject": "Alert: {{message}}",
  "body": "<h1>{{title}}</h1><p>{{content}}</p>",
  "content_type": "HTML",
  "cc": "cc@example.com",
  "bcc": ""
}
```

**Using Manual Access Token:**
```json
{
  "auth_method": "token",
  "access_token": "your_access_token_here",
  "to": "recipient@example.com",
  "subject": "Alert: {{message}}",
  "body": "<h1>{{title}}</h1><p>{{content}}</p>",
  "content_type": "HTML",
  "cc": "cc@example.com",
  "bcc": ""
}
```

**Options:**
- `auth_method`: "oauth" (recommended) or "token"
- `client_id`: Azure AD application client ID (for OAuth)
- `client_secret`: Azure AD application client secret (for OAuth)
- `tenant_id`: Azure AD directory tenant ID (for OAuth)
- `access_token`: Microsoft Graph API access token (for token method)
- `to`: Recipient email address (supports liquid templating)
- `subject`: Email subject (supports liquid templating)
- `body`: Email body (supports liquid templating and HTML)
- `content_type`: "HTML" or "Text" (default: "HTML")
- `cc`: CC recipients (optional, supports liquid templating)
- `bcc`: BCC recipients (optional, supports liquid templating)

## Example Workflows

### Email Alert System
1. **OutlookAgent** monitors inbox for emails with specific criteria
2. **TriggerAgent** filters emails based on content
3. **OutlookSenderAgent** sends alerts to different recipients

### Email Forwarding
1. **OutlookAgent** receives emails from specific senders
2. **DataTransformationAgent** modifies the content
3. **OutlookSenderAgent** forwards to different recipients

### Digest Emails
1. Multiple agents collect data throughout the day
2. **TriggerAgent** creates a daily digest
3. **OutlookSenderAgent** sends summary email

## Security Considerations

- Store access tokens securely (consider using Huginn's credentials system)
- Use the minimum required API permissions
- Regularly rotate access tokens
- Consider using Microsoft's recommended authentication flows
- Validate email content to prevent injection attacks

## Development

Running `rake` will clone and set up Huginn in `spec/huginn` to run the specs of the Gem in Huginn as if they would be build-in Agents. The desired Huginn repository and branch can be modified in the `Rakefile`:

```ruby
HuginnAgent.load_tasks(branch: '<your branch>', remote: 'https://github.com/<github user>/huginn.git')
```

Make sure to delete the `spec/huginn` directory and re-run `rake` after changing the `remote` to update the Huginn source code.

After the setup is done `rake spec` will only run the tests, without cloning the Huginn source again.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release` to create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

1. Fork it ( https://github.com/[my-github-username]/huginn_outlook_agent/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request

## License

This project is licensed under the MIT License - see the LICENSE.txt file for details.
