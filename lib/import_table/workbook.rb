require 'import_table/cell'
require 'import_table/sheet'
require 'import_table/options'
require 'import_table/row'
require 'import_table/setting'
require 'import_table/symbolize'

module ImportTable
  class Workbook
    include Cell
    include Options
    include Sheet
    include Row
    include Setting
    include Symbolize

    attr_reader :options, :settings, :info, :uniques

    # @param file [String,StringIO]:
    # @param options [Hash]:
    # @attribute
    # * extension [Symbol]: - :xls, :xlsx, :ods, :csv
    # * csv_options: {col_sep: "\t"}
    # * default_sheet: [Integer|String]
    def initialize(file, options = {})
      @options = symbolize(options)

      @file    = file
      @uniques = {}
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
      @settings = settings || {}
      change_sheet(@settings[:sheet])

      review_settings(:read)

      range = @settings[:first_row].upto(@settings[:last_row])

      block ? rows_streaming(range, &block) : rows(range)
    end

    # Read N || 10 rows for preview
    # @param settings [Hash]:
    #    sheet [Integer|String] - number or name of sheet
    #    first_row_preview [Integer] - first reading row (default 2);
    #    last_row_preview [Integer] - last reading row (default 11).
    # @return [Array] - rows.
    def preview(settings = {})
      @settings = settings || {}
      change_sheet(@settings[:sheet])
      review_settings(:preview)
      range = @settings[:first_row_preview].upto(@settings[:last_row_preview])

      rows(range)
    end

    private

    def open
      @workbook = Roo::Spreadsheet.open(@file, @options)

      info!
    end
  end
end
