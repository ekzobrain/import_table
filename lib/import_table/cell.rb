module ImportTable
  module Cell
    private

    def prepare_cell(row, param, line)
      return nil unless param[:column]

      result = cell_to_type(row[param[:column]], param)
      result = regexp_cell(result, param) if param[:regexp_search]
      check_unique(result, param, line)

      result
    end

    def regexp_cell(value, param)
      if param[:regexp_type] == :sub
        value.sub(param[:regexp_search], param[:regexp_replace])
      else
        value.gsub(param[:regexp_search], param[:regexp_replace])
      end
    end

    def cell_to_type(value, param)
      case param[:type]
      when :string
        t_string(value, param)
      when :boolean
        t_boolean(value)
      when :integer
        t_integer(value)
      when :float
        t_float(value)
      when :date
        t_date(value, param, param[:type])
      else
        value
      end
    end

    def check_unique(value, param, line)
      name = param[:unique]
      return unless name

      @uniques[name][:values].include?(value) ? add_not_unique(value, name, line) : @uniques[name][:values] << value
    end

    def add_not_unique(value, name, line)
      not_uniq = @uniques[name][:not_unique]

      not_uniq.include?(value) ? not_uniq[value] << line : not_uniq[value] = [line]

      @uniques[name][:not_unique_count] += 1
    end

    def t_string(value, param)
      return String(value) unless param[:format]

      t_date(value, param, param[:type]) if %i[date date_time].include? param[:format]
    end

    def t_date(value, param, type)
      date = type == :date ? Date.parse(value) : DateTime.parse(value)

      date&.strftime(param[:strftime])
    end

    def t_boolean(value)
      return true if value&.match?(/^(true|t|yes|y|1)$/i)

      false if value&.empty? || value&.match?(/^(false|f|no|n|0)$/i)
    end

    def t_integer(value)
      Integer(value)
    end

    def t_float(value)
      Float(value)
    end
  end
end
