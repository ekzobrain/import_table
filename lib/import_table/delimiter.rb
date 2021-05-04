module ImportTable
  class Delimiter
    class << self
      DELIMITERS = [';', ':', '|', "\'\t\'"].freeze

      def type(file)
        fr = first_row(file)
        p fr
        p DELIMITERS.reduce({}) { |res, deli| res.merge deli => fr.count(deli) }
        # .sort { |a, b| b[1] <=> a[1] }
      end

      private

      def first_row(file)
        if file.instance_of?(StringIO)
          file.first
        else
          File.open(file, &:readline)
        end
      end
    end
  end
end
