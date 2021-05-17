require_relative '../spec_helper'
# require 'rspec/context/private'

TEST_SHEETS_INFO = {
  default_sheet: nil, sheets_count: 2, sheets_name: %w[Sheet1 Sheet2], sheet_current: 'Sheet1', sheets: {
    'Sheet1' => {
      first_row: 1, last_row: 10, first_column: 1, last_column: 8, first_column_literal: 'A', last_column_literal: 'H'
    },
    'Sheet2' => {
      first_row: 1, last_row: 4, first_column: 1, last_column: 8, first_column_literal: 'A', last_column_literal: 'H'
    }
  }
}.freeze

TEST_SHEETS_DEFAULT = {
  default_sheet: 'Sheet2', sheets_count: 2, sheets_name: %w[Sheet1 Sheet2], sheet_current: 'Sheet2', sheets: {
    'Sheet2' => {
      first_row: 1, last_row: 4, first_column: 1, last_column: 8, first_column_literal: 'A', last_column_literal: 'H'
    }
  }
}.freeze

# rubocop:disable Metrics/BlockLength
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
  end

  describe '.preview' do
    it 'Change current sheet by sheet name' do
      xls_w2s.preview(current_sheet: 'Sheet2')

      expect(xls_w2s.info).to include(sheet_current: 'Sheet2')
    end

    it 'Change current sheet by sheet index' do
      xls_w2s.preview(current_sheet: 2)

      expect(xls_w2s.info).to include(sheet_current: 'Sheet2')
    end

    it 'Read xls row 2 ' do
      rows = xls_w2s.preview

      expect(rows[1][2 .. 4]).to eq(['Hashimoto', 'Female', 'Great Britain'])
    end

    it 'Read csv row 2 ' do
      rows = csv_eo_wd.preview

      expect(rows[2][2 .. 4]).to eq(%w[Gent Männlich Frankreich])
    end

    it 'Verify last_row' do
      rows = csv_eo_wd.preview(last_row: 112).last

      expect(rows).to eq(['9', 'Vincenza', 'Weiland', 'Weiblich', 'Vereinigte Staaten', '40', '21/05/2015', '6548'])
    end
  end
end
# rubocop:enable Metrics/BlockLength
