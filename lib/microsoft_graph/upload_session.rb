class MicrosoftGraph
  class UploadSession < Resource
    attr_reader :message

    class << self
      def create_path(upload_session_hash = {})
        message = upload_session_hash["message"]
        "/me/messages/#{message.id}/attachments/createUploadSession"
      end

      def associations
        %w(message)
      end
    end

    def upload_url
      @hash["uploadUrl"]
    end
  end
end
