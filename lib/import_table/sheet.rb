module ImportTable
  module Sheet
    private

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
    # @param name [String|Integer]
    # @return [String] - valid name
    def verify_sheet_name(name)
      if name.is_a?(Integer)
        raise SheetNotFound, "Sheet index '#{name}' out of range" unless @info[:sheets_count] - 1 >= name

        @workbook.sheets[name]
      else
        unless @info[:sheets_name].include?(name)
          raise SheetNotFound, "Sheet name '#{name}' not in list #{@info[:sheets_name]}"
        end

        name
      end
    end

    def verify_max_row(row)
      last_row = current_last_row

      row > last_row ? last_row : row
    end

    def current_last_row
      @info[:sheets][@info[:sheet_current]][:last_row]
    end
  end
end
