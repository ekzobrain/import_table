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
    def initialize(file)
      @file = file

      mime  = FileMagic.mime
      @mime =
        if file.instance_of?(StringIO)
          mime&.io(file, 1024, true)&.split('; ')
        else
          mime&.file(file)&.split('; ')
        end
    end

    def type?
      MIME_TYPES[@mime[0]] if @mime&.any?
    end

    def encoding?
      @mime[1]&.sub('charset=', '') if @mime&.any?
    end
  end
end
