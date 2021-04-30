module ImportTable
  class Workbook
    # @param [Hash] options
    ## :extension - :xls, :xlsx, :csv,
    ## :csv_options: {col_sep: "\t"}
    def initialize(file, options = {})
      # begin
      @file = Roo::Spreadsheet.open(file, options)
      # rescue
      # "error"
      # end
    end

    # @return [Hash]
    def sheets_info
      return nil unless @file

      {
        sheets_count:  @file.sheets.count,
        sheets_name:   @file.sheets,
        sheet_current: @file.default_sheet
      }
    end

    # @return [Int]
    def last_row
      return nil unless @file

      @file.last_row
    end
  end
end
