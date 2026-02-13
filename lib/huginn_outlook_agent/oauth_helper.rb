require 'oauth2'
require 'json'
require 'time'

module Agents
  class OAuthHelper
    attr_reader :client_id, :client_secret, :tenant_id, :access_token, :refresh_token, :expires_at

    def initialize(client_id, client_secret, tenant_id, refresh_token = nil)
      @client_id = client_id
      @client_secret = client_secret
      @tenant_id = tenant_id
      @refresh_token = refresh_token
      @access_token = nil
      @expires_at = nil
    end

    def valid_access_token?
      @access_token && @expires_at && Time.now < @expires_at
    end

    def get_access_token
      return @access_token if valid_access_token?
      
      if @refresh_token
        refresh_access_token
      else
        acquire_new_token
      end
    end

    private

    def acquire_new_token
      # Check for nil or empty values
      if !@client_id || @client_id.empty?
        raise "Client ID is nil or empty"
      end
      
      if !@client_secret || @client_secret.empty?
        raise "Client Secret is nil or empty"
      end
      
      if !@tenant_id || @tenant_id.empty?
        raise "Tenant ID is nil or empty"
      end
      
      # Use HTTParty directly for OAuth2 client credentials flow
      token_url = "https://login.microsoftonline.com/" + @tenant_id.to_s + "/oauth2/v2.0/token"
      
      headers = {
        'Content-Type' => 'application/x-www-form-urlencoded'
      }
      
      body = URI.encode_www_form({
        grant_type: 'client_credentials',
        client_id: @client_id,
        client_secret: @client_secret,
        scope: 'https://graph.microsoft.com/.default'
      })

      begin
        response = HTTParty.post(token_url, body: body, headers: headers)
        
        unless response.success?
          raise "OAuth2 Error: #{response.code} - #{response.body}"
        end
        
        data = JSON.parse(response.body)
        store_token_response(data)
        data['access_token']
      rescue JSON::ParserError => e
        raise "Failed to parse OAuth response: #{e.message}"
      rescue => e
        raise "Failed to acquire access token: #{e.message} (#{e.class})"
      end
    end

    def refresh_access_token
      # Only attempt refresh if we have a refresh token
      return acquire_new_token unless @refresh_token && !@refresh_token.empty?
      
      # Build URLs safely without interpolation that might conflict with Liquid
      token_url = "https://login.microsoftonline.com/" + @tenant_id.to_s + "/oauth2/v2.0/token"
      
      client = OAuth2::Client.new(
        @client_id,
        @client_secret,
        site: "https://login.microsoftonline.com",
        token_url: token_url
      )

      begin
        response = OAuth2::AccessToken.new(
          client,
          @access_token,
          refresh_token: @refresh_token
        ).refresh!

        store_token(response)
        response.token
      rescue OAuth2::Error => e
        # If refresh fails, try to acquire new token
        acquire_new_token
      rescue => e
        # If refresh fails, try to acquire new token
        acquire_new_token
      end
    end

    def store_token_response(data)
      @access_token = data['access_token']
      @refresh_token = data['refresh_token'] if data['refresh_token']
      @expires_at = data['expires_in'] ? Time.now + data['expires_in'].to_i : Time.now + 3600
    end

    def store_token(response)
      # Handle both OAuth2 gem response and direct HTTP response
      if response.respond_to?(:token)
        # OAuth2 gem response object
        @access_token = response.token
        @refresh_token = response.refresh_token if response.respond_to?(:refresh_token) && response.refresh_token
        @expires_at = response.expires_at ? Time.parse(response.expires_at.to_s) : Time.now + 3600
      else
        # Direct HTTP response hash
        store_token_response(response)
      end
    end
  end
end
