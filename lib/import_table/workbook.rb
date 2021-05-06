module ImportTable
  class Workbook
    attr_reader :options, :settings, :sheets_info

    # @param [Hash] options
    # extension: - :xls, :xlsx, :ods, :csv
    # csv_options: {col_sep: "\t"}
    def initialize(file, options = {})
      @options = options
      @file    = file

      review_options
      open
    end

    def read(settings = {})
      change_current_sheet(settings)

      settings[:first_row].upto(settings[:last_row]) { |line| yield @workbook.row(line) }
    end

    #  Sets the first and last line for preview
    # (current: first = 1, last = 7 )
    # @param settings [Hash]:
    #    first_row;
    #    last_row.
    def preview(settings = {})
      change_current_sheet(settings)
      settings[:first_row] = 2 unless settings.include?(:first_row)
      settings[:last_row]  = 7 unless settings.include?(:last_row)

      settings[:first_row].upto(settings[:last_row]).map { |line| @workbook.row(line) }
    end

    # @return [Int]
    def last_row
      @workbook.last_row
    end

    def info
      @workbook.info
    end

    private

    def open
      case @options[:extension]
      when :csv
        @workbook = Roo::CSV.new(@file, @options)
      when :xls
        @workbook = Roo::Spreadsheet.open(@file, @options)
      end

      sheets_info!
    end

    def change_current_sheet(settings)
      return unless settings.include?(:current_sheet)

      new_current = settings.delete(:current_sheet)

      if new_current.is_a?(Integer)
        @workbook.default_sheet = @workbook.sheets[new_current] if @sheets_info[:sheets_count] - 1 >= new_current
      elsif sheets_info[:sheets_name].include?(new_current)
        @workbook.default_sheet = new_current
      end

      sheets_info!
    end

    # @return [Hash]
    def sheets_info!
      @sheets_info = {
        sheets_count:  @workbook.sheets.count,
        sheets_name:   @workbook.sheets,
        sheet_current: @workbook.default_sheet
      }
    end

    def review_options
      if @options.empty?
        make_options
      elsif @options[:csv_options]&.include?(:col_sep) && @options[:extension] != :csv
        @options.merge!(extension: :csv)
      end
    end

    def make_options
      mime = ImportTable::Mime.new(@file)

      @options[:extension] = mime.type?
      @options.merge!(csv_options: { col_sep: mime.delimiter? }) if mime.delimiter?
    end
  end
end
