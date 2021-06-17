class MicrosoftGraph
  class ResourceCollection < Resource
    include Enumerable

    def resources
      @resources ||= hash["value"].map do |resource_hash|
        resource_class.new(microsoft_graph, resource_hash)
      end
    end

    def first
      resources.first
    end

    def resource_class
      self.class.name.sub("Collection", "").constantize
    end

    def each_page
      current_page = self
      loop do
        yield current_page
        break if current_page.last_page?
        current_page = current_page.next_page
      end
    end

    def each
      resources.each { |resource| yield resource }
    end

    def count
      hash["value"].count
    end

    def empty?
      count == 0
    end

    def last_page?
      next_page_url.nil?
    end

    def next_page
      response = microsoft_graph.get_url(next_page_url)
      response_hash = JSON.parse(response.body)
      MessageCollection.new(microsoft_graph, response_hash)
    end

    def next_page_url
      hash["@odata.nextLink"]
    end
  end
end
