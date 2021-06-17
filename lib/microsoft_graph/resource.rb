class MicrosoftGraph
  class Resource
    class << self
      def find(microsoft_graph, id)
        new(microsoft_graph, "id" => id).retrieve
      end

      def create(microsoft_graph, resource_hash)
        create_path = if method(:create_path).arity == 0
          create_path()
        else
          create_path(resource_hash)
        end
        response = microsoft_graph.post(create_path, resource_hash.except(*associations))
        resource_hash = JSON.parse(response.body)
        new(microsoft_graph, resource_hash)
      end

      def associations
        []
      end
    end

    attr_reader :microsoft_graph, :hash

    def initialize(microsoft_graph, hash)
      @microsoft_graph = microsoft_graph
      @hash = hash
    end

    def id
      hash["id"]
    end

    def retrieve
      response = microsoft_graph.get(path)
      response_hash = JSON.parse(response.body)
      self.class.new(microsoft_graph, response_hash)
    end

    def update(resource_hash)
      response = microsoft_graph.patch(
        path,
        resource_hash
      )
      response_hash = JSON.parse(response.body)
      self.class.new(microsoft_graph, response_hash)
    end

    def delete
      microsoft_graph.delete(path)
    end
  end
end
