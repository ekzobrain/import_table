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

      prepare_mapping if @settings.include?(:mapping)
    end

    def verify_rows_settings(for_method = :read)
      first, last, last_row =
        for_method == :read ? [:first_row, :last_row, current_last_row] : [:first_row_preview, :last_row_preview, 11]

      @settings[first] = @settings[first] ? verify_max_row(@settings[first]) : verify_max_row(2)
      @settings[last]  = @settings[last] ? verify_max_row(@settings[last]) : verify_max_row(last_row)
    end

    def prepare_mapping
      @settings[:mapping_type] = @settings[:mapping_type].to_sym if @settings[:mapping_type].is_a?(String)
      @settings[:mapping_type] = :hash unless %i[hash array].include?(@settings[:mapping_type])
      @settings[:mapping].each do |name, params|
        letter_to_number(params)
        param_to_sym(params)
        default_strftime(params)
        unique_set(name, params)
      end
    end

    def letter_to_number(params)
      params[:column] = ::Roo::Utils.letter_to_number(params[:column].to_s) - 1 unless params[:column].is_a?(Integer)
    end

    def param_to_sym(params)
      %i[type format].each { |i| params[i] = params[i].to_sym if params[i].is_a?(String) }
    end

    def default_strftime(params)
      DEFAULT_STRFTIME.each do |format, value|
        params[:strftime] = value if params[:format] == format && params[:strftime].nil?
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
  end
end
