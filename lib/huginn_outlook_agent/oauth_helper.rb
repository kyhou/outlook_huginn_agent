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
      
      # Build URLs safely without interpolation that might conflict with Liquid
      token_url = "https://login.microsoftonline.com/" + @tenant_id.to_s + "/oauth2/v2.0/token"
      authorize_url = "https://login.microsoftonline.com/" + @tenant_id.to_s + "/oauth2/v2.0/authorize"
      
      client = OAuth2::Client.new(
        @client_id,
        @client_secret,
        site: "https://login.microsoftonline.com",
        token_url: token_url,
        authorize_url: authorize_url
      )

      begin
        response = client.client_credentials.get_token(
          scope: 'https://graph.microsoft.com/.default'
        )
        
        store_token(response)
        response.token
      rescue OAuth2::Error => e
        # Parse the error response for more details
        error_details = ""
        if e.response
          if e.response.body
            begin
              error_json = JSON.parse(e.response.body)
              error_details = " - #{error_json['error_description'] || error_json['error'] || ''}"
            rescue JSON::ParserError
              error_details = " - #{e.response.body}"
            end
          else
            error_details = " - No response body"
          end
        else
          error_details = " - No response object"
        end
        
        raise "OAuth2 Error: #{e.message}#{error_details} - Check client_id, client_secret, and tenant_id"
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

    def store_token(response)
      @access_token = response.token
      @refresh_token = response.refresh_token if response.respond_to?(:refresh_token) && response.refresh_token
      @expires_at = response.expires_at ? Time.parse(response.expires_at.to_s) : Time.now + 3600
    end
  end
end
