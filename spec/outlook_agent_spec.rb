require 'rails_helper'
require 'huginn_agent/spec_helper'

describe Agents::OutlookAgent do
  before(:each) do
    @valid_options = {
      'mode' => 'receive',
      'access_token' => 'test_token',
      'folder' => 'inbox',
      'since' => '',
      'mark_as_read' => false
    }
    @checker = Agents::OutlookAgent.new(:name => "OutlookAgent", :options => @valid_options)
    @checker.user = users(:bob)
    @checker.save!
  end

  describe "validation" do
    it "should be valid with default options" do
      expect(@checker).to be_valid
    end

    it "should require mode" do
      @checker.options['mode'] = ''
      expect(@checker).not_to be_valid
      expect(@checker.errors[:base]).to include("Mode is required")
    end

    it "should require access token" do
      @checker.options['access_token'] = ''
      expect(@checker).not_to be_valid
      expect(@checker.errors[:base]).to include("Access token is required")
    end

    it "should validate folder for receive mode" do
      @checker.options['folder'] = 'invalid'
      expect(@checker).not_to be_valid
      expect(@checker.errors[:base]).to include("Folder must be 'inbox', 'sent', 'drafts', or 'deleted'")
    end

    it "should validate send mode requirements" do
      @checker.options['mode'] = 'send'
      @checker.options['to'] = ''
      @checker.options['subject'] = ''
      @checker.options['body'] = ''
      expect(@checker).not_to be_valid
      expect(@checker.errors[:base]).to include("Recipient (to) is required for send mode")
      expect(@checker.errors[:base]).to include("Subject is required for send mode")
      expect(@checker.errors[:base]).to include("Body is required for send mode")
    end

    it "should validate content type" do
      @checker.options['content_type'] = 'invalid'
      expect(@checker).not_to be_valid
      expect(@checker.errors[:base]).to include("Content type must be 'HTML' or 'Text'")
    end
  end

  describe "#working?" do
    it "should return true when checked without error" do
      allow(@checker).to receive(:checked_without_error?).and_return(true)
      expect(@checker.working?).to be_truthy
    end

    it "should return true when received event without error" do
      allow(@checker).to receive(:received_event_without_error?).and_return(true)
      expect(@checker.working?).to be_truthy
    end
  end

  describe "#check" do
    it "should call receive_emails when in receive mode" do
      @checker.options['mode'] = 'receive'
      expect(@checker).to receive(:receive_emails)
      @checker.check
    end

    it "should not call receive_emails when in send mode" do
      @checker.options['mode'] = 'send'
      expect(@checker).not_to receive(:receive_emails)
      @checker.check
    end
  end

  describe "#receive" do
    let(:event) { Event.new(payload: { 'message' => 'test' }) }

    it "should call send_email for each event when in send mode" do
      @checker.options['mode'] = 'send'
      expect(@checker).to receive(:send_email).with(event)
      @checker.receive([event])
    end

    it "should not call send_email when in receive mode" do
      @checker.options['mode'] = 'receive'
      expect(@checker).not_to receive(:send_email)
      @checker.receive([event])
    end
  end
end

describe Agents::OutlookSenderAgent do
  before(:each) do
    @valid_options = {
      'access_token' => 'test_token',
      'to' => 'test@example.com',
      'subject' => 'Test Subject',
      'body' => 'Test Body',
      'content_type' => 'HTML'
    }
    @sender = Agents::OutlookSenderAgent.new(:name => "OutlookSenderAgent", :options => @valid_options)
    @sender.user = users(:bob)
    @sender.save!
  end

  describe "validation" do
    it "should be valid with default options" do
      expect(@sender).to be_valid
    end

    it "should require access token" do
      @sender.options['access_token'] = ''
      expect(@sender).not_to be_valid
      expect(@sender.errors[:base]).to include("Access token is required")
    end

    it "should require recipient" do
      @sender.options['to'] = ''
      expect(@sender).not_to be_valid
      expect(@sender.errors[:base]).to include("Recipient (to) is required")
    end

    it "should require subject" do
      @sender.options['subject'] = ''
      expect(@sender).not_to be_valid
      expect(@sender.errors[:base]).to include("Subject is required")
    end

    it "should require body" do
      @sender.options['body'] = ''
      expect(@sender).not_to be_valid
      expect(@sender.errors[:base]).to include("Body is required")
    end
  end

  describe "#working?" do
    it "should return true when received event without error" do
      allow(@sender).to receive(:received_event_without_error?).and_return(true)
      expect(@sender.working?).to be_truthy
    end
  end

  describe "#receive" do
    let(:event) { Event.new(payload: { 'message' => 'test' }) }

    it "should call send_email for each event" do
      expect(@sender).to receive(:send_email).with(event)
      @sender.receive([event])
    end
  end
end
