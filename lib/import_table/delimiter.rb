module ImportTable
  class Delimiter
    DELIMITERS = [';', ':', '|', "\'\t\'"].freeze

    class << self
      # Define the delimiter used in csv
      # @param [String|StringIO]
      # @return [Nil|String]
      def type(file)
        fr = first_row(file)

        delim = DELIMITERS.reduce({}) { |res, deli| res.merge deli => fr.count(deli) }.max_by(&:last)
        delim.last.zero? ? nil : delim.first.sub("\'\t\'", "\t")
      end

      private

      # Read the first line in csv
      # @param [String|StringIO]
      # @return [File]
      def first_row(file)
        file.instance_of?(StringIO) ? file.first : File.open(file, &:readline)
      end
    end
  end
end
