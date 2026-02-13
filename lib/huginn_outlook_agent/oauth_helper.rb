require 'oauth2'
require 'json'
require 'time'

module HuginnOutlookAgent
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
      # Debug: Log credential values (without exposing secrets)
      puts "DEBUG: client_id present: #{@client_id.present?}"
      puts "DEBUG: client_secret present: #{@client_secret.present?}"
      puts "DEBUG: tenant_id present: #{@tenant_id.present?}"
      puts "DEBUG: tenant_id value: '#{@tenant_id}'"
      
      # Build URLs safely without interpolation that might conflict with Liquid
      token_url = "https://login.microsoftonline.com/" + @tenant_id.to_s + "/oauth2/v2.0/token"
      authorize_url = "https://login.microsoftonline.com/" + @tenant_id.to_s + "/oauth2/v2.0/authorize"
      
      puts "DEBUG: token_url: #{token_url}"
      
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
        puts "DEBUG: OAuth2::Error details: #{e.class} - #{e.message}"
        raise "OAuth2 Error: #{e.message} - Check client_id, client_secret, and tenant_id"
      rescue => e
        puts "DEBUG: General error details: #{e.class} - #{e.message}"
        raise "Failed to acquire access token: #{e.message} (#{e.class})"
      end
    end

    def refresh_access_token
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
        raise "OAuth2 Error during refresh: #{e.message}"
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
