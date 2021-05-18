require 'import_table/sheet'
require 'import_table/options'

module ImportTable
  class Workbook
    include Options
    include Sheet

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

    # Read N || 10 rows for preview
    # @param settings [Hash]:
    #    sheet [Integer|String] - number or name of sheet
    #    first_row [Integer] - first reading row (default 2);
    #    last_row [Integer] - last reading row (default 12).
    #    mapping_type [Symbol] - :Hash or :Array
    #    mapping [Hash] - schema
    #    -
    # @return [Array|Yield] - rows.
    def read(settings = {}, &block)
      @settings = settings
      change_sheet(settings[:sheet])
      verify_rows_settings(:read)
      prepare_mapping if @settings.include?(:mapping)

      if block
        rows_streaming(@settings[:first_row], @settings[:last_row], &block)
      else
        rows(@settings[:first_row], @settings[:last_row])
      end
    end

    # Read N || 10 rows for preview
    # @param settings [Hash]:
    #    sheet [Integer|String] - number or name of sheet
    #    first_row_preview [Integer] - first reading row (default 2);
    #    last_row_preview [Integer] - last reading row (default 11).
    # @return [Array] - rows.
    def preview(settings = {})
      @settings = settings
      change_sheet(settings[:sheet])
      verify_rows_settings(:preview)

      rows(@settings[:first_row_preview], @settings[:last_row_preview])
    end

    private

    def open
      @workbook = Roo::Spreadsheet.open(@file, @options)

      info!
    end

    def rows(first_row, last_row)
      if @settings.include?(:mapping)
        if @settings[:mapping_type] == :hash
          first_row.upto(last_row).map { |line| prepare_row_hash(@workbook.row(line), @settings[:mapping]) }
        end
      else
        first_row.upto(last_row).map { |line| @workbook.row(line) }
      end
    end

    def rows_streaming(first_row, last_row)
      if @settings.include?(:mapping)
        if @settings[:mapping_type] == :hash
          first_row.upto(last_row).each { |line| yield prepare_row_hash(@workbook.row(line), @settings[:mapping]) }
        end
      else
        first_row.upto(last_row).each { |line| yield @workbook.row(line) }
      end
    end

    def prepare_row_hash(row, mapping)
      mapping.transform_values { |param| prepare_cell(row, param) }
    end

    def prepare_cell(row, param)
      return nil unless param[:column]

      cell_to_type(row[param[:column]], param[:type])
    end

    def cell_to_type(value, type)
      case type
      when 'string'
        value
      when 'boolean'
        bool value
      when 'integer'
        Integer(value) if value
      when 'float'
        Float(value)
      when 'date'
        Date.parse(value)
      else
        value
      end
    end

    def prepare_mapping
      @settings[:mapping_type] = :hash unless @settings.include?(:mapping_type)

      @settings[:mapping].each do |_, params|
        params[:column] = ::Roo::Utils.letter_to_number(params[:column].to_s) - 1 unless params[:column].is_a?(Integer)
      end
    end

    def verify_rows_settings(for_method = :read)
      first, last, last_row =
        for_method == :read ? [:first_row, :last_row, current_last_row] : [:first_row_preview, :last_row_preview, 11]

      @settings[first] = @settings[first] ? verify_max_row(@settings[first]) : verify_max_row(2)
      @settings[last]  = @settings[last] ? verify_max_row(@settings[last]) : verify_max_row(last_row)
    end

    def bool(value)
      return true if value&.match?(/^(true|t|yes|y|1)$/i)
      return false if value&.empty? || value&.match?(/^(false|f|no|n|0)$/i)

      nil
    end
  end
end
