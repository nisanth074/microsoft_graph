class MicrosoftGraph
  class Subscription < Resource
    class << self
      def create_path
        "/subscriptions"
      end
    end

    def renew
      new_expiration_time = Time.zone.now + 4230.minutes - 1.minute # Microsoft allows any new expiration time until 4230 minutes from now
      microsoft_graph.patch(path, "expirationDateTime" => new_expiration_time)
    end

    def path
      "/subscriptions/#{id}"
    end
  end
end
