module ImportTable
  class Workbook
    attr_reader :options, :settings, :info

    # @param file [String|StringIO]:
    # @param options [Hash]:
    # @attribute:
    # * extension [Symbol]: - :xls, :xlsx, :ods, :csv
    # * csv_options: {col_sep: "\t"}
    # * default_sheet: [Integer|String]
    def initialize(file, options = {})
      @options = options.slice(:extension, :csv_options, :default_sheet)
      @file    = file

      review_options
      open
    end

    def read(settings = {})
      change_sheet(settings)

      settings[:first_row].upto(settings[:last_row]) { |line| yield @workbook.row(line) }
    end

    #  Sets the first and last line for preview
    # (current: first = 1, last = 7 )
    # @param settings [Hash]:
    #    first_row;
    #    last_row.
    def preview(settings = {})
      change_sheet(settings.delete(:current_sheet))

      first_row = settings[:first_row] ? verify_max_row(settings[:first_row]) : verify_max_row(2)
      last_row  = settings[:last_row] ? verify_max_row(settings[:last_row]) : verify_max_row(10)

      first_row.upto(last_row).map { |line| @workbook.row(line) }
    end

    private

    def open
      @workbook =
        case @options[:extension]
        when :csv
          Roo::CSV.new(@file, @options)
        when :xls
          Roo::Spreadsheet.open(@file, @options)
        when :xlsx
          Roo::Excelx.new(@file, @options)
        end

      info!
    end

    # Take information about sheets.
    def info!
      @info.merge!(
        sheets_count:  @workbook.sheets.count,
        sheets_name:   @workbook.sheets,
        sheet_current: @workbook.default_sheet,
        sheets:        {}
      )

      change_sheet(@info[:default_sheet])
      sheets_info
    end

    # Collects information about sheets.
    # If set default_sheet, collect only for default.
    def sheets_info
      if @info[:default_sheet]
        @info[:default_sheet] = verify_sheet_name(@info[:default_sheet])
        @info[:sheets].merge!(sheet_info(@info[:default_sheet]))
      else
        @info[:sheets_name].each { |name| @info[:sheets].merge!(sheet_info(name)) }
        @workbook.default_sheet = @info[:sheet_current]
      end
    end

    def sheet_info(name)
      @workbook.default_sheet = name
      {
        name => {
          first_row:            @workbook.first_row,
          last_row:             @workbook.last_row,
          first_column:         @workbook.first_column,
          last_column:          @workbook.last_column,
          first_column_literal: ::Roo::Utils.number_to_letter(@workbook.first_column),
          last_column_literal:  ::Roo::Utils.number_to_letter(@workbook.last_column)
        }
      }
    end

    def change_sheet(name)
      return unless name

      @info[:sheet_current] = @workbook.default_sheet = verify_sheet_name(name)
    end

    # Checks the name or number of a sheet in the sheets list.
    # Return valid name.
    # @param name [String|Integer]
    # @return [String]
    def verify_sheet_name(name)
      if name.is_a?(Integer)
        p name
        raise SheetNotFound, "Sheet index '#{name}' out of range" unless @info[:sheets_count] >= name

        @workbook.sheets[name - 1]
      else
        unless @info[:sheets_name].include?(name)
          raise SheetNotFound, "Sheet name '#{name}' not in list #{@info[:sheets_name]}"
        end

        name
      end
    end

    def verify_max_row(row, lag = 0)
      last_row = @info[:sheets][@info[:sheet_current]][:last_row]
      row > last_row ? last_row + lag : row
    end

    # Checks options for reading a file.
    def review_options
      @info = { default_sheet: @options.delete(:default_sheet) }

      if @options.empty?
        create_options
      elsif @options[:csv_options]&.include?(:col_sep) && @options[:extension] != :csv
        @options.merge!(extension: :csv)
      end
    end

    # Creates parameters for reading the file.
    # Detect file type and delimiter for csv files.
    def create_options
      mime = ImportTable::Mime.new(@file)

      @options[:extension] = mime.type?
      @options.merge!(csv_options: { col_sep: mime.delimiter? }) if mime.delimiter?
    end
  end
end
