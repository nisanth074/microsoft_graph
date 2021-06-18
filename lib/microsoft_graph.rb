require "microsoft_graph/version"

require "microsoft_graph/resource"
require "microsoft_graph/resource_collection"

require "microsoft_graph/upload_session"
require "microsoft_graph/message"
require "microsoft_graph/subscription"

require "microsoft_graph/message_collection"

class MicrosoftGraph
  attr_reader :oauth_authorization

  def initialize(oauth_authorization)
    @oauth_authorization = oauth_authorization
  end

  def get(path, options = {})
    url = base_url + path
    get_url(url, options)
  end

  def get_url(url, options = {})
    Hooligan.get(url, { headers: authorization_headers }.merge(options))
  end

  def post(path, body_hash = nil, options = {})
    url = base_url + path
    options = { headers: post_headers }
    if body_hash.nil?
      options[:headers]["Content-Type"] = "application/json"
    else
      options[:headers]["Content-Type"] = "application/json"
      options[:body] = body_hash.to_json
    end

    Hooligan.post(url, options)
  end

  def patch(path, body_hash = {})
    url = base_url + path
    options = { headers: authorization_headers }
    options[:headers]["Content-Type"] = "application/json"
    options[:body] = body_hash.to_json

    Hooligan.patch(url, options)
  end

  def delete(path)
    url = base_url + path
    Hooligan.delete(url, headers: authorization_headers)
  end

  def most_recent_message
    Message.most_recent(self)
  end

  def messages_received_after(time)
    Message.received_after(self, time)
  end

  # @todo Add rspecs for this method
  def find_message_by_message_id(message_id)
    Message.find_by_message_id(self, message_id)
  end

  def create_message(message_hash)
    Message.create(self, message_hash)
  end

  def send_message(message_hash)
    Message._send(self, message_hash)
  end

  def send_message_with_file_attachments(message_hash, file_attachment_hashes)
    message = Message.create(microsoft_graph, message_hash)
    message.add_file_attachments(file_attachment_hashes)
  end

  def create_reply(parent_message_id, message_hash)
    Message.new(self, "id" => parent_message_id).create_reply(message_hash)
  end

  def send_reply(parent_message_id, message_hash)
    Message.new(self, id: parent_message_id).send_reply(message_hash)
  end

  def create_subscription(subscription_hash)
    Subscription.create(self, subscription_hash)
  end

  private

  def post_headers
    # By default, Outlook changes the ID of a resource, like a message, when the resource is moved to a different folder.
    # Prevent this behaviour by supplying the +Prefer: IdType="ImmutableId"+ header.
    #
    # @see https://docs.microsoft.com/en-us/graph/outlook-immutable-id
    # @see https://developer.microsoft.com/en-us/outlook/blogs/announcing-immutable-id-for-outlook-resources-in-microsoft-graph/
    authorization_headers.merge("Prefer" => 'IdType="ImmutableId"')
  end

  def authorization_headers
    {
      "Authorization" => "Bearer #{access_token}"
    }
  end

  def access_token
    @oauth_authorization.latest_access_token
  end

  def base_url
    "https://graph.microsoft.com/v1.0"
  end
end
