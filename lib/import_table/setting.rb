module ImportTable
  module Setting
    DEFAULT_STRFTIME = {
      date:      '%Y-%m-%d',
      date_time: '%Y-%m-%dT%H:%M:%SZ'
    }.freeze

    private

    def review_settings(for_method = :read)
      verify_rows_settings(for_method)
      @settings = symbolize(@settings)

      return unless @settings.include?(:mapping)

      extract_array_mapping if @settings[:mapping].is_a?(Array)
      prepare_mapping
    end

    def extract_array_mapping
      @settings[:mapping] =
        @settings[:mapping].map do |params|
          key = params.delete(:to) || params.delete(:title) || params[:column]
          [key.to_sym, params]
        end.to_h
    end

    def verify_rows_settings(for_method = :read)
      first, last, last_row =
        for_method == :read ? [:first_row, :last_row, current_last_row] : [:first_row_preview, :last_row_preview, 11]

      @settings[first] = @settings[first] ? verify_max_row(@settings[first]) : verify_max_row(2)
      @settings[last]  = @settings[last] ? verify_max_row(@settings[last]) : verify_max_row(last_row)
    end

    def prepare_mapping
      return_type_set

      @settings[:mapping].each do |name, params|
        letter_to_number(params)
        param_to_sym(params)
        default_strftime(params)
        unique_set(name, params)
        regexp_set(params)
      end
    end

    def return_type_set
      @settings[:mapping_type] = @settings[:return_type] if @settings[:return_type]
      @settings[:mapping_type] = @settings[:mapping_type].to_sym if @settings[:mapping_type].is_a?(String)
      @settings[:mapping_type] = :hash unless %i[hash array].include?(@settings[:mapping_type])
    end

    def letter_to_number(params)
      params[:column] = ::Roo::Utils.letter_to_number(params[:column].to_s) - 1 unless params[:column].is_a?(Integer)
    end

    def param_to_sym(params)
      %i[type format].each { |i| params[i] = params[i].to_sym if params[i].is_a?(String) }
    end

    def default_strftime(params)
      DEFAULT_STRFTIME.each do |format, value|
        params[:strftime] = value if (params[:format] == format || params[:type] == format) && params[:strftime].nil?
      end
    end

    def unique_set(name, params)
      return unless params[:unique]

      params[:unique] = name

      @uniques[name] = {
        column:           params[:column],
        not_unique:       {},
        not_unique_count: 0,
        values:           []
      }
    end

    def regexp_set(params)
      return unless params[:regexp_search]

      validate_regexp(params)
      validate_replace(params)
      value_to_sym(params, :regexp_type)
      params[:regexp_type] = :sub unless params[:regexp_type] && %i[sub gsub].include?(params[:regexp_type])
    end

    def validate_regexp(params)
      params[:regexp_search] = Regexp.new(params[:regexp_search])
    rescue RegexpError => e
      raise RegexpError, "Invalid regular expression: #{e.message}"
    end

    def validate_replace(params)
      return unless params[:regexp_search]

      raise RegexpError, 'No replacement variable.' unless params[:regexp_replace]
    end

    def value_to_sym(params, params_name)
      [params_name].each { |i| params[i] = params[i].to_sym if params[i].is_a?(String) }
    end
  end
end
