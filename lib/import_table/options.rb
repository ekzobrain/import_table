module ImportTable
  module Options
    FILE_TYPES = %i[csv csv xls xlsx ods].freeze

    private

    # Checks options for reading a file.
    def review_options
      @info                = { default_sheet: @options.delete(:default_sheet) }
      @options[:extension] = :csv if @options[:csv_options]&.include?(:col_sep) && @options[:extension] != :csv

      check_extension
      check_delimiter
    end

    def check_extension
      if @file.instance_of?(StringIO)
        raise MissingRequiredOption, 'extension' unless @options.include?(:extension)
      else
        @options[:extension] = Roo::Spreadsheet.extension_for(@file, @options)
      end
      raise UnsupportedFileType, @options[:extension] unless FILE_TYPES.include?(@options[:extension])
    end

    def check_delimiter
      return unless @options[:extension] == :csv
      return if @options[:csv_options]&.include?(:col_sep)

      delim = ImportTable::Delimiter.type(@file)
      return unless delim

      @options[:csv_options]           = {} unless @options.include?(:csv_options)
      @options[:csv_options][:col_sep] = delim
    end
  end
end
