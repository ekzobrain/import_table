module ImportTable
  module Symbolize
    def symbolize(value)
      case value
      when Hash
        symbolize_recursive(value)
      when Array
        value.map { |v| symbolize(v) }
      else
        value
      end
    end

    def symbolize_recursive(hash)
      {}.tap { |h| hash.each { |key, value| h[key.to_sym] = symbolize(value) } }
    end
  end
end
