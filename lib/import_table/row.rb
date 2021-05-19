module ImportTable
  module Row
    private

    # Read rows
    def rows(range)
      return range.map { |line| @workbook.row(line) } unless @settings.include?(:mapping)

      if @settings[:mapping_type] == :hash
        range.map { |line| prepare_row_hash(@workbook.row(line), @settings[:mapping], line) }
      else
        range.map { |line| prepare_row_array(@workbook.row(line), @settings[:mapping], line) }
      end
    end

    # Read rows streaming
    def rows_streaming(range)
      return range.each { |line| yield @workbook.row(line) } unless @settings.include?(:mapping)

      if @settings[:mapping_type] == :hash
        range.each { |line| yield prepare_row_hash(@workbook.row(line), @settings[:mapping], line) }
      else
        range.each { |line| yield prepare_row_array(@workbook.row(line), @settings[:mapping], line) }
      end
    end

    def prepare_row_hash(row, mapping, line)
      mapping.transform_values { |params| prepare_cell(row, params, line) }
    end

    def prepare_row_array(row, mapping, line)
      mapping.map { |_, params| prepare_cell(row, params, line) }
    end
  end
end
