require 'httparty'
require 'json'
require 'time'
require_relative 'oauth_helper'

module Agents
  class OutlookAgent < Agent
    cannot_receive_events!
    can_dry_run!
    default_schedule 'every_5m'

    description <<-MD
      The Outlook Agent integrates with Microsoft Outlook via the Microsoft Graph API to:
      
      - **Receive emails**: Monitor inbox for new emails and create events
      - **Send emails**: Send emails through Outlook based on incoming events
      
      ## Authentication
      
      This agent uses OAuth 2.0 with Microsoft Graph API. You need to:
      
      1. Register an application in Azure Active Directory
      2. Configure the required API permissions (Mail.Read, Mail.Send)
      3. Obtain a client ID, client secret, and tenant ID
      4. Get an access token (or use the provided OAuth flow)
      
      ## Configuration
      
      **For receiving emails:**
      - `mode`: 'receive' 
      - `access_token`: Microsoft Graph API access token
      - `folder`: Inbox folder to monitor (default: 'inbox')
      - `since`: Filter emails received since this time (ISO 8601 format)
      - `mark_as_read`: Whether to mark processed emails as read (default: false)
      
      **For sending emails:**
      - `mode`: 'send'
      - `access_token`: Microsoft Graph API access token
      - `to`: Recipient email address (can use liquid templating)
      - `subject`: Email subject (can use liquid templating)
      - `body`: Email body (can use liquid templating, supports HTML)
      - `content_type`: 'HTML' or 'Text' (default: 'HTML')
      
      The agent will interpolate liquid templates in the `to`, `subject`, and `body` fields
      using the payload of incoming events when in 'send' mode.
    MD

    def default_options
      {
        'mode' => 'receive',
        'auth_method' => 'oauth',
        'client_id' => '',
        'client_secret' => '',
        'tenant_id' => '',
        'access_token' => '',
        'refresh_token' => '',
        'folder' => 'inbox',
        'since' => '',
        'mark_as_read' => false,
        'to' => '',
        'subject' => '',
        'body' => '',
        'content_type' => 'HTML'
      }
    end

    def validate_options
      errors.add(:base, "Mode is required") unless options['mode'].present?
      
      if options['auth_method'] == 'oauth'
        errors.add(:base, "Client ID is required for OAuth") unless options['client_id'].present?
        errors.add(:base, "Client Secret is required for OAuth") unless options['client_secret'].present?
        errors.add(:base, "Tenant ID is required for OAuth") unless options['tenant_id'].present?
      else
        errors.add(:base, "Access token is required") unless options['access_token'].present?
      end
      
      if options['mode'] == 'send'
        errors.add(:base, "Recipient (to) is required for send mode") unless options['to'].present?
        errors.add(:base, "Subject is required for send mode") unless options['subject'].present?
        errors.add(:base, "Body is required for send mode") unless options['body'].present?
      end
      
      if options['mode'] == 'receive'
        errors.add(:base, "Folder must be 'inbox', 'sent', 'drafts', or 'deleted'") unless ['inbox', 'sent', 'drafts', 'deleted'].include?(options['folder'])
      end
      
      if options['content_type'].present?
        errors.add(:base, "Content type must be 'HTML' or 'Text'") unless ['HTML', 'Text'].include?(options['content_type'])
      end
    end

    def working?
      checked_without_error? || received_event_without_error?
    end

    def check
      if options['mode'] == 'receive'
        receive_emails
      end
    end

    def receive(incoming_events)
      if options['mode'] == 'send'
        incoming_events.each do |event|
          send_email(event)
        end
      end
    end

    private

    def oauth_helper
      @oauth_helper ||= begin
        if options['auth_method'] == 'oauth'
          Agents::OAuthHelper.new(
            interpolated['client_id'] || '',
            interpolated['client_secret'] || '', 
            interpolated['tenant_id'] || '',
            interpolated['refresh_token'] || ''
          )
        end
      end
    end

    def current_access_token
      if options['auth_method'] == 'oauth' && oauth_helper
        oauth_helper.get_access_token
      else
        options['access_token']
      end
    end

    def graph_api_url
      'https://graph.microsoft.com/v1.0'
    end

    def headers
      {
        'Authorization' => "Bearer #{current_access_token}",
        'Content-Type' => 'application/json'
      }
    end

    def receive_emails
      # Client credentials flow cannot access user mailboxes - need to inform user
      error("Client credentials authentication cannot access user mailboxes. Please use delegated authentication with user login instead of application permissions.")
      return
    end

    def send_email(event)
      interpolated = interpolate_options(event.payload)
      
      payload = {
        'message' => {
          'subject' => interpolated['subject'],
          'body' => {
            'contentType' => interpolated['content_type'] || 'HTML',
            'content' => interpolated['body']
          },
          'toRecipients' => format_recipients(interpolated['to'])
        }
      }
      
      if interpolated['cc'].present?
        payload['message']['ccRecipients'] = format_recipients(interpolated['cc'])
      end
      
      if interpolated['bcc'].present?
        payload['message']['bccRecipients'] = format_recipients(interpolated['bcc'])
      end
      
      url = "#{graph_api_url}/me/sendMail"
      response = HTTParty.post(url, body: payload.to_json, headers: headers)
      
      unless response.success?
        error("Failed to send email: #{response.code} - #{response.message}")
        return
      end
      
      log("Email sent successfully to #{interpolated['to']}")
    end

    def mark_as_read(email_id)
      url = "#{graph_api_url}/me/messages/#{email_id}"
      payload = { 'isRead' => true }
      
      HTTParty.patch(url, body: payload.to_json, headers: headers)
    end

    def extract_email_address(field)
      return nil unless field
      field['emailAddress']['address']
    end

    def extract_body(email)
      body = email['body']
      return '' unless body
      
      if body['contentType'] == 'text'
        body['content']
      else
        body['content']
      end
    end

    def format_recipients(recipients)
      return [] unless recipients.present?
      
      Array(recipients).map do |recipient|
        {
          'emailAddress' => {
            'address' => recipient.strip
          }
        }
      end
    end
  end
end
