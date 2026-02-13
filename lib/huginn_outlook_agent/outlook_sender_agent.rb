require 'httparty'
require 'json'
require_relative 'oauth_helper'

module HuginnOutlookAgent
  class OutlookSenderAgent < Agent
    can_dry_run!
    default_schedule 'never'

    description <<-MD
      The Outlook Sender Agent sends emails through Microsoft Outlook via Microsoft Graph API.
      
      This agent is specifically designed for sending emails and should be triggered by other agents.
      
      ## Configuration
      
      - `access_token`: Microsoft Graph API access token
      - `to`: Recipient email address (can use liquid templating)
      - `subject`: Email subject (can use liquid templating)
      - `body`: Email body (can use liquid templating, supports HTML)
      - `content_type`: 'HTML' or 'Text' (default: 'HTML')
      - `cc`: CC recipients (optional, can use liquid templating)
      - `bcc`: BCC recipients (optional, can use liquid templating)
      
      The agent will interpolate liquid templates in all fields using the payload of incoming events.
      
      ## Usage
      
      1. Set up Microsoft Graph API access token
      2. Configure the email parameters
      3. Trigger this agent from other agents using the "Send to webhook" action
    MD

    def default_options
      {
        'auth_method' => 'oauth',
        'client_id' => '',
        'client_secret' => '',
        'tenant_id' => '',
        'access_token' => '',
        'refresh_token' => '',
        'to' => '',
        'subject' => '',
        'body' => '',
        'content_type' => 'HTML',
        'cc' => '',
        'bcc' => ''
      }
    end

    def validate_options
      if options['auth_method'] == 'oauth'
        errors.add(:base, "Client ID is required for OAuth") unless options['client_id'].present?
        errors.add(:base, "Client Secret is required for OAuth") unless options['client_secret'].present?
        errors.add(:base, "Tenant ID is required for OAuth") unless options['tenant_id'].present?
      else
        errors.add(:base, "Access token is required") unless options['access_token'].present?
      end
      
      errors.add(:base, "Recipient (to) is required") unless options['to'].present?
      errors.add(:base, "Subject is required") unless options['subject'].present?
      errors.add(:base, "Body is required") unless options['body'].present?
      
      if options['content_type'].present?
        errors.add(:base, "Content type must be 'HTML' or 'Text'") unless ['HTML', 'Text'].include?(options['content_type'])
      end
    end

    def working?
      received_event_without_error?
    end

    def receive(incoming_events)
      incoming_events.each do |event|
        send_email(event)
      end
    end

    private

    def oauth_helper
      @oauth_helper ||= begin
        if options['auth_method'] == 'oauth'
          interpolated = interpolate_options({})
          HuginnOutlookAgent::OAuthHelper.new(
            interpolated['client_id'],
            interpolated['client_secret'], 
            interpolated['tenant_id'],
            interpolated['refresh_token']
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

    def interpolate_options(payload)
      interpolated = {}
      
      %w[to subject body cc bcc content_type].each do |field|
        value = options[field]
        interpolated[field] = value.present? ? liquid_interpolate(value, payload) : ''
      end
      
      interpolated
    end
  end
end
