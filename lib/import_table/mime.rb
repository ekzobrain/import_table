require 'ruby-filemagic'

module ImportTable
  class Mime
    MIME_TYPES = {
      'application/csv'                                                   => :csv,
      'text/plain'                                                        => :csv,
      'application/vnd.ms-excel'                                          => :xls,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' => :xlsx,
      'application/vnd.oasis.opendocument.spreadsheet'                    => :ods
    }.freeze

    # @param [String|StringIO]
    # @return @mime [Array]
    def initialize(file)
      @file = file

      mime  = FileMagic.mime
      @mime = file.instance_of?(StringIO) ? mime.io(file, 1024, true) : mime.file(file)

      raise ImportTable::FileCannotOpen, @mime if @mime&.include?('cannot open')

      @mime = @mime.split('; ')
      mime.close
    end

    # Defining file type
    def type?
      raise ImportTable::UnsupportedType, @mime[0] unless MIME_TYPES.include?(@mime[0])

      MIME_TYPES[@mime[0]]
    end

    # Defining field separator
    # @return [String,Nil]
    def delimiter?
      case @mime[0]
      when 'application/csv'
        ','
      when 'text/plain'
        ImportTable::Delimiter.type(@file)
      end
    end

    # Defining file encoding
    # @return [String|Nil]
    def encoding?
      @mime[1]&.sub('charset=', '')
    end
  end
end
