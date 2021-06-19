class MicrosoftGraph
  class Message < Resource
    class << self
      def create_path
        "/me/messages"
      end

      def most_recent(microsoft_graph)
        query_params = {
          "$orderBy" => "receivedDateTime desc",
          "$top" => 1
        }
        path = "/me/messages"
        response = microsoft_graph.get(path, query: query_params)
        response_hash = JSON.parse(response.body)
        MessageCollection.new(microsoft_graph, response_hash).first
      end

      def received_after(microsoft_graph, time)
        query_params = {
          "$orderBy" => "receivedDateTime asc",
          "$filter" => "receivedDateTime gt #{time.iso8601}"
        }
        path = "/me/messages"
        response = microsoft_graph.get(path, query: query_params)
        response_hash = JSON.parse(response.body)
        MessageCollection.new(microsoft_graph, response_hash)
      end

      # @todo Add rspecs for this method
      def find_by_message_id(microsoft_graph, message_id)
        query_params = {
          "$filter" => "internetMessageId eq '<#{message_id}>'",
          "$top" => 1
        }
        path = "/me/messages"
        response = microsoft_graph.get(path, query: query_params)
        response_hash = JSON.parse(response.body)
        MessageCollection.new(microsoft_graph, response_hash)
      end

      def _send(microsoft_graph, message_hash)
        draft_message = Message.create(microsoft_graph, message_hash)
        draft_message._send
      end

      def send_reply(reply_message_hash)
        reply_message = create_reply(reply_message_hash)
        reply_message._send
      end
    end

    def path
      "/me/messages/#{id}"
    end

    def subject
      hash["subject"]
    end

    def sender_email
      hash["sender"]["emailAddress"]["address"]
    end

    def received_at
      Time.parse(hash["receivedDateTime"])
    end

    def draft?
      !!hash["isDraft"]
    end

    def raw
      response = microsoft_graph.get("/me/messages/#{id}/$value")
     response.body
    end

    def draft?
      !!hash["isDraft"]
    end

    # @note Please that only a draft message can be sent.
    #   A draft message is a message whose +isDraft` property is true.
    def _send
      response = microsoft_graph.post("/me/messages/#{id}/send")
      true
    end

    def add_file_attachment(file_attachment_hash)
      file_content = Base64.decode64(file_attachment_hash["contentBytes"])
      file_size = file_content.bytesize

      if file_size < 3.megabytes
        add_small_file_attachment(file_attachment_hash)
      else
        add_large_file_attachment(file_attachment_hash)
      end
    end

    def add_small_file_attachment(file_attachment_hash)
      file_attachment_hash = file_attachment_hash.merge("@odata.type" => "#microsoft.graph.fileAttachment")
      create_attachment_path = "#{path}/attachments"
      microsoft_graph.post(create_attachment_path, file_attachment_hash)
    end

    def add_large_file_attachment(file_attachment_hash)
      file_content = Base64.decode64(file_attachment_hash["contentBytes"])
      file_size = file_content.bytesize

      # Create an upload session
      upload_session_hash = {
        "attachmentItem" => {
          "attachmentType" => "file",
          "name" => file_attachment_hash["name"],
          "size" => file_size
        },
        "message" => self
      }
      upload_session_hash["attachmentItem"].merge!(
        "isInline" => true,
        "contentId" => file_attachment_hash["contentId"]
      ) if file_attachment_hash["isInline"]
      upload_session = UploadSession.create(microsoft_graph, upload_session_hash)

      # Upload file in 5 megabyte slices

      uploadUrl = upload_session.upload_url
      slice_size = 5.megabytes
      final_request_response = (0..(file_size - 1)).step(slice_size).each do |index|
        from = index
        slice = file_content.byteslice(from, index + slice_size - 1)
        to = index + [slice.bytesize, slice_size].min - 1
        headers = {
          "Content-Type" => "application/octet-stream",
          "Content-Length" => slice.bytesize.to_s,
          "Content-Range" => "bytes #{from}-#{to}/#{file_size}"
        }
        Hooligan.put(uploadUrl, headers: headers, body: slice)
      end
      # @todo Raise error unless the final request response status is 201
    end

    def add_file_attachments(file_attachment_hashes)
      file_attachment_hashes.each { |file_attachment_hash| add_file_attachment(file_attachment_hash) }
    end

    def create_reply(reply_message_hash = {})
      create_reply_path = "/me/messages/#{id}/createReply"
      response = microsoft_graph.post(create_reply_path)
      response_hash = JSON.parse(response.body)
      message = Message.new(microsoft_graph, response_hash)
      return message if reply_message_hash.empty?
      message.update(reply_message_hash)
    end
  end
end
