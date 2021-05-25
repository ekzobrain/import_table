require_relative '../spec_helper'
# require 'rspec/context/private'

TEST_SHEETS_INFO = {
  default_sheet: nil, sheets_count: 2, sheets_name: %w[Sheet1 Sheet2], sheet_current: 'Sheet1', sheets: {
    'Sheet1' => {
      first_row: 1, last_row: 10, first_column: 0, last_column: 7, first_column_literal: 'A', last_column_literal: 'H'
    },
    'Sheet2' => {
      first_row: 1, last_row: 4, first_column: 0, last_column: 7, first_column_literal: 'A', last_column_literal: 'H'
    }
  }
}.freeze

TEST_SHEETS_DEFAULT = {
  default_sheet: 'Sheet2', sheets_count: 2, sheets_name: %w[Sheet1 Sheet2], sheet_current: 'Sheet2', sheets: {
    'Sheet2' => {
      first_row: 1, last_row: 4, first_column: 0, last_column: 7, first_column_literal: 'A', last_column_literal: 'H'
    }
  }
}.freeze

# rubocop:disable Metrics/BlockLength
# rubocop:disable RSpec/MultipleExpectations
describe ImportTable::Workbook do
  # xls with two sheets
  let(:xls_w2s) { described_class.new(get_file('file_example_XLS_10_2s.xls')) }
  # CSV without options & with delimiter
  let(:csv_eo_wd) { described_class.new(get_file('file_example_CSV_10_comma_de.csv')) }

  describe '.new' do
    let(:f_xls2s) { get_file('file_example_XLS_10_2s.xls') }

    it 'Creates an instance' do
      expect(csv_eo_wd).to be_a(described_class)
    end

    it 'Review options for File 1 - empty options for csv with delimiter' do
      wb = csv_eo_wd

      expect(wb.options).to eq({ csv_options: { col_sep: ',' }, file_warning: :ignore })
    end

    it 'Review options for File 2 - with options csv_options: {col_sep:}' do
      wb = described_class.new(get_file('file_example_TSV_10.tsv'), { csv_options: { col_sep: "\t" } })

      expect(wb.options).to include({ csv_options: { col_sep: "\t" } })
    end

    it 'Review options for File 3 - empty options for csv without delimiter' do
      wb = described_class.new(get_file('file_example_CSV_10_without_delimiter.csv'))

      expect(wb.info[:sheets]['default']).to include({ last_row: 10, last_column_literal: 'A' })
    end

    it 'Review options for File 4 - with options for unsupported file' do
      expect { described_class.new(get_file('file_example_RTF_100kB.rtf')) }
        .to raise_error(ImportTable::UnsupportedFileType)
    end

    it 'Review options for StringIO 1 - with options' do
      wb = described_class.new(get_string_io('file_example_CSV_10_semicolon.csv'), { extension: :csv })

      expect(wb.options).to eq({ csv_options: { col_sep: ';' }, file_warning: :ignore })
    end

    it 'Review options for StringIO 2 - without options' do
      expect { described_class.new(get_string_io('file_example_CSV_10_semicolon.csv')) }
        .to raise_error(ImportTable::MissingRequiredOption)
    end

    it 'Review options for StringIO 3 - with options for unsupported file' do
      expect { described_class.new(get_string_io('file_example_CSV_10_semicolon.csv'), extension: :png) }
        .to raise_error(ImportTable::UnsupportedFileType)
    end

    it 'Sheets info 1 - full info' do
      expect(xls_w2s.info).to eq(TEST_SHEETS_INFO)
    end

    it 'Sheets info 2 - only default sheet info (by number)' do
      wb = described_class.new(f_xls2s, { default_sheet: 2 })

      expect(wb.info).to eq(TEST_SHEETS_DEFAULT)
    end

    it 'Sheets info 3 - only default sheet info (by name)' do
      wb = described_class.new(f_xls2s, { default_sheet: 'Sheet2' })

      expect(wb.info).to eq(TEST_SHEETS_DEFAULT)
    end

    it 'Sheets info 4 - out of range' do
      expect { described_class.new(f_xls2s, { default_sheet: 3 }) }.to raise_error(ImportTable::SheetNotFound)
    end

    it 'Sheets info 5 - not in list' do
      expect { described_class.new(f_xls2s, { default_sheet: 'ts' }) }.to raise_error(ImportTable::SheetNotFound)
    end

    it 'Symbolize options 1 - without csv_options' do
      options = { 'extension' => 'csv' }
      wb      = described_class.new(get_string_io('file_example_CSV_10_semicolon.csv'), options)

      expect(wb.options).to eq({ csv_options: { col_sep: ';' }, file_warning: :ignore })
    end

    it 'Symbolize options 2 - with csv_options' do
      options = { 'extension' => 'csv', 'csv_options' => { 'col_sep' => ';' } }
      wb      = described_class.new(get_string_io('file_example_CSV_10_semicolon.csv'), options)

      expect(wb.options).to eq({ csv_options: { col_sep: ';' }, file_warning: :ignore })
    end
  end

  describe '.preview' do
    it 'Change current sheet by sheet name' do
      xls_w2s.preview(sheet: 'Sheet2')

      expect(xls_w2s.info).to include(sheet_current: 'Sheet2')
    end

    it 'Change current sheet by sheet index' do
      xls_w2s.preview(sheet: 2)

      expect(xls_w2s.info).to include(sheet_current: 'Sheet2')
    end

    it 'Read xls row 2 ' do
      rows = xls_w2s.preview

      expect(rows[1][2 .. 4]).to eq(['Hashimoto', 'Female', 'Great Britain'])
    end

    it 'Gen header' do
      rows = xls_w2s.preview({ return_type: :hash })

      expect(rows[2])
        .to eq({ A: 3.0, B: 'Philip', C: 'Gent', D: 'Male', E: 'France', F: 36.0, G: '21/05/2015', H: 2587.0 })
    end

    it 'Without - Gen header' do
      rows = xls_w2s.preview({ return_type: :array })

      expect(rows[2])
        .to eq([3.0, 'Philip', 'Gent', 'Male', 'France', 36.0, '21/05/2015', 2587.0])
    end

    it 'Read csv row 2 ' do
      rows = csv_eo_wd.preview

      expect(rows[2][2 .. 4]).to eq(%w[Gent Männlich Frankreich])
    end

    it 'Check last_row' do
      rows = csv_eo_wd.preview(last_row: 112).last

      expect(rows).to eq(['9', 'Vincenza', 'Weiland', 'Weiblich', 'Vereinigte Staaten', '40', '21/05/2015', '6548'])
    end
  end

  describe '.read' do
    let(:expected_rows) do
      [
        [1.0, 'Dulce', 'Abril', 'Female', 'United States', 32.0, '15/10/2017', 1562.0],
        [2.0, 'Mara', 'Hashimoto', 'Female', 'Great Britain', 25.0, '16/08/2016', 1582.0],
        [3.0, 'Philip', 'Gent', 'Male', 'France', 36.0, '21/05/2015', 2587.0],
        [4.0, 'Kathleen', 'Hanner', 'Female', 'United States', 25.0, '15/10/2017', 3549.0],
        [5.0, 'Nereida', 'Magwood', 'Female', 'United States', 58.0, '16/08/2016', 2468.0],
        [6.0, 'Gaston', 'Brumm', 'Male', 'United States', 24.0, '21/05/2015', 2554.0],
        [7.0, 'Etta', 'Hurn', 'Female', 'Great Britain', 56.0, '15/10/2017', 3598.0],
        [8.0, 'Earlean', 'Melgar', 'Female', 'United States', 27.0, '16/08/2016', 2456.0],
        [9.0, 'Vincenza', 'Weiland', 'Female', 'United States', 40.0, '21/05/2015', 6548.0]
      ]
    end

    let(:unique_test_mapping) do
      {
        Country:  { column: 'E', type: :string, unique: true },
        LastName: { column: 'C', type: :string, unique: true }
      }
    end

    let(:reg_fn) { { column: 'B', regexp_search: '^(.).*', regexp_replace: '\1.' } }

    it 'Read 1 - without mapping' do
      expect(xls_w2s.read).to eq(expected_rows)
    end

    # 'Type :integer'
    it 'Read 2.1 - result - hash; A & F to integer - without format' do
      mapping = { Index: { column: :A, type: :integer }, Age: { column: :F, type: :integer } }
      rows    = xls_w2s.read(mapping_type: :hash, mapping: mapping)

      expect(rows).to eq(expected_rows.map { |item| { Index: Integer(item[0]), Age: Integer(item[5]) } })
    end

    it 'Read 2.2 (streaming) - with mapping: result - hash; A & F to integer - without format' do
      i = 0
      xls_w2s.read(mapping_type: :hash, mapping: { Age: { column: :F, type: :integer } }) do |row|
        expect(row).to eq({ Age: Integer(expected_rows[i][5]) })
        i += 1
      end
    end

    # 'Float'
    it 'Read 2.3 - result - hash; A & F to float' do
      mapping = { Index: { column: :A, type: :float }, Age: { column: :F, type: :float } }
      rows    = xls_w2s.read(mapping_type: :hash, mapping: mapping)

      expect(rows).to eq(expected_rows.map { |item| { Index: Float(item[0]), Age: Float(item[5]) } })
    end

    # 'Boolean'
    it 'Read 2.4 - result - hash; b to boolean' do
      wb   = described_class.new(get_file('file_example_CSV_10_comma_bool.csv'))
      rows = wb.read(return_type: :array, mapping: [{ column: :B, type: :boolean }])

      expect(rows).to eq([[true], [false], [true], [false], [true], [false], [true], [false], [true], [false]])
    end

    # 'Date'
    it 'Read 2.5.1 - col G to type date, default strftime (RFC 3339, section 5.6)' do
      mapping = { Date: { column: 6, type: :date } }
      rows    = xls_w2s.read(mapping_type: :hash, mapping: mapping)

      expect(rows).to eq(expected_rows.map { |item| { Date: Date.parse(item[6]).strftime('%Y-%m-%d') } })
    end

    it 'Read 2.5.2 - col G to type date, strftime manual (Y.m.d)' do
      mapping = { Date: { column: 6, type: :date, strftime: '%Y.%m.%d' } }
      rows    = xls_w2s.read(mapping_type: :hash, mapping: mapping)

      expect(rows).to eq(expected_rows.map { |item| { Date: Date.parse(item[6]).strftime('%Y.%m.%d') } })
    end

    it 'Read 2.5.3 - col G to type date_time, default strftime (RFC 3339, section 5.6)' do
      mapping = { Date: { column: 6, type: :date_time } }
      rows    = xls_w2s.read(mapping_type: :hash, mapping: mapping)

      expect(rows).to eq(expected_rows.map { |item| { Date: Date.parse(item[6]).strftime('%Y-%m-%dT%H:%M:%SZ') } })
    end

    it 'Read 2.5.4 - col G to type date_time, strftime manual (Y-m-d H:M:S)' do
      mapping = { Date: { column: 6, type: :date_time, strftime: '%Y-%m-%d %H:%M:%S' } }
      rows    = xls_w2s.read(mapping_type: :hash, mapping: mapping)

      expect(rows).to eq(expected_rows.map { |item| { Date: Date.parse(item[6]).strftime('%Y-%m-%d %H:%M:%S') } })
    end

    # 'Type :string'
    # A to String, G to String without format
    it 'Read 3.1.1 - result - array; A & G to string - without format' do
      mapping = [{ column: :A, type: :string }, { column: :G, type: :string }]
      rows    = xls_w2s.read(return_type: :array, mapping: mapping)

      expect(rows).to eq(expected_rows.map { |item| [String(item[0]), String(item[6])] })
    end

    # 'Type :string & format: :date || dateTime'
    # G to String, format: Date, strftime - RFC 3339, section 5.6 (default)
    it 'Read 3.2.1 - col G to string, format: date, with default strftime' do
      mapping = { Date: { column: 6, type: :string, format: 'date' } }
      rows    = xls_w2s.read(mapping_type: :hash, mapping: mapping)

      expect(rows).to eq(expected_rows.map { |item| { Date: Date.parse(item[6]).strftime('%Y-%m-%d') } })
    end

    # G to String, format: Date, strftime - '%Y.%m.%d'
    it 'Read 3.2.2 - col G to string, format: date, with strftime manual' do
      mapping = { Date: { column: 6, type: :string, format: :date, strftime: '%Y.%m.%d' } }
      rows    = xls_w2s.read(mapping_type: :hash, mapping: mapping)

      expect(rows).to eq(expected_rows.map { |item| { Date: Date.parse(item[6]).strftime('%Y.%m.%d') } })
    end

    # G to String, format: DateTime, strftime - RFC 3339, section 5.6 (default)
    it 'Read 3.2.3 - col G to string, format: date_time, with default strftime' do
      mapping = { Date: { column: 6, type: :string, format: 'date_time' } }
      rows    = xls_w2s.read(mapping_type: :hash, mapping: mapping)

      expect(rows).to eq(expected_rows.map { |item| { Date: Date.parse(item[6]).strftime('%Y-%m-%dT%H:%M:%SZ') } })
    end

    # G to String, format: DateTime, strftime - %Y-%m-%d %H:%M:%S
    it 'Read 3.2.4 - col G to string, format: date_time, with strftime manual' do
      mapping = { Date: { column: 6, type: 'string', format: 'date_time', strftime: '%Y-%m-%d %H:%M:%S' } }
      rows    = xls_w2s.read(mapping_type: :hash, mapping: mapping)

      expect(rows).to eq(expected_rows.map { |item| { Date: Date.parse(item[6]).strftime('%Y-%m-%d %H:%M:%S') } })
    end

    # 'Param Unique'
    it 'Read 4.1 - col C to string, without not_unique cell' do
      xls_w2s.read(mapping_type: :hash, mapping: unique_test_mapping)

      expect(xls_w2s.uniques[:LastName][:not_unique]).to eq({})
      expect(xls_w2s.uniques[:LastName][:not_unique_count]).to eq(0)
      expect(xls_w2s.uniques[:LastName][:column]).to eq(2)
    end

    it 'Read 4.2 - col E to string, with not_unique cell' do
      xls_w2s.read(mapping_type: :array, mapping: unique_test_mapping)

      expect(xls_w2s.uniques[:Country][:not_unique])
        .to eq({ 'Great Britain' => [8], 'United States' => [5, 6, 7, 9, 10] })

      expect(xls_w2s.uniques[:Country][:not_unique_count]).to eq(6)
    end

    it 'Read 4.3 (streaming) - E to string, with not_unique cell' do
      i = 0
      xls_w2s.read(mapping_type: :array, mapping: unique_test_mapping) do
        expect(xls_w2s.uniques[:Country][:not_unique_count]).to eq([0, 0, 0, 1, 2, 3, 4, 5, 6][i])
        i += 1
      end
    end

    # 'Param Regexp'
    it 'Read 5.1.1 - Settings - Invalid regular expression' do
      mapping = { mapping: { FN: { column: 'B', regexp_search: '?', regexp_replace: '\1.' } } }

      expect { xls_w2s.read(mapping) }.to raise_error(RegexpError)
    end

    it 'Read 5.1.2 - Settings - No replacement variable' do
      mapping = { mapping: { FN: { column: 'B', regexp_search: '?' } } }

      expect { xls_w2s.read(mapping) }.to raise_error(RegexpError)
    end

    it 'Read 5.2.1 - set regexp_type :sub (default)' do
      xls_w2s.read(mapping: { FN: reg_fn })

      expect(xls_w2s.settings[:mapping][:FN][:regexp_type]).to eq(:sub)
    end

    it 'Read 5.2.2 - set regexp_type :gsub (change value type - from string)' do
      xls_w2s.read(mapping: { FN: reg_fn.merge!(regexp_type: 'gsub') })

      expect(xls_w2s.settings[:mapping][:FN][:regexp_type]).to eq(:gsub)
    end

    it 'Read 5.2.3 - set regexp_type :sub (default) if wrong type' do
      xls_w2s.read(mapping: { FN: reg_fn.merge!(regexp_type: :zub) })

      expect(xls_w2s.settings[:mapping][:FN][:regexp_type]).to eq(:sub)
    end

    it 'Read 5.3.1 - replace with regexp_type :sub ' do
      rows = xls_w2s.read(mapping_type: :array, mapping: { FN: reg_fn })

      expect(rows).to eq([['D.'], ['M.'], ['P.'], ['K.'], ['N.'], ['G.'], ['E.'], ['E.'], ['V.']])
    end

    it 'Read 5.3.2 - replace with regexp_type :gsub ' do
      mapping = { FN: { column: 'B', regexp_search: '([ie])', regexp_replace: '\1.', regexp_type: :gsub } }
      xls_w2s
      rows = xls_w2s.read(mapping_type: :array, mapping: mapping)[2 .. 4]

      expect(rows).to eq([['Phi.li.p'], ['Kathle.e.n'], ['Ne.re.i.da']])
    end

    # 'Settings Symbolize'
    it 'Read 6.1 Symbolize for mapping_type => array' do
      xls_w2s.read('mapping_type' => 'array',
                   'mapping'      => { 'Country' => { 'column' => 'E', 'type' => 'string', 'unique' => true } })

      expect(xls_w2s.uniques[:Country][:not_unique])
        .to eq({ 'Great Britain' => [8], 'United States' => [5, 6, 7, 9, 10] })
      expect(xls_w2s.uniques[:Country][:not_unique_count]).to eq(6)
    end

    it 'Read 6.2 Symbolize convert array in mapping' do
      xls_w2s.read('mapping_type' => 'array',
                   'mapping'      => [{ 'to' => 'Country', 'column' => 'E' }])
      expect(xls_w2s.settings[:mapping]).to eq({ Country: { column: 4 } })
    end
  end
end
# rubocop:enable RSpec/MultipleExpectations
# rubocop:enable Metrics/BlockLength
