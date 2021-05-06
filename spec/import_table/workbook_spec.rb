require_relative '../spec_helper'
# require 'rspec/context/private'

# rubocop:disable Metrics/BlockLength
describe ImportTable::Workbook do
  # xls with two sheets
  let(:xls_w2s) { described_class.new(get_file('file_example_XLS_10_2s.xls')) }
  # CSV without options & with delimiter
  let(:csv_eo_wd) { described_class.new(get_file('file_example_CSV_10_comma_de.csv')) }

  describe '.new' do
    it 'Creates an instance' do
      expect(csv_eo_wd).to be_a(described_class)
    end

    it 'Review options 1 - empty options for csv with delimiter' do
      wb = csv_eo_wd

      expect(wb.options).to eq({ extension: :csv, csv_options: { col_sep: ',' } })
    end

    it 'Review options 2 - empty options for csv without delimiter' do
      wb = described_class.new(get_file('file_example_CSV_10_without_delimiter.csv'))

      expect(wb.options).to eq({ extension: :csv })
    end

    it 'Review options 3 - with options csv_options: {col_sep:}' do
      wb = described_class.new(get_file('file_example_TSV_10.tsv'), { csv_options: { col_sep: "\t" } })

      expect(wb.options).to include({ extension: :csv, csv_options: { col_sep: "\t" } })
    end

    it 'Take sheets info' do
      sheets_info = { sheets_count: 2, sheets_name: %w[Sheet1 Sheet2], sheet_current: 'Sheet1' }
      expect(xls_w2s.sheets_info).to include(sheets_info)
    end
  end

  describe '.preview' do
    it 'Change current sheet by sheet name' do
      xls_w2s.preview(current_sheet: 'Sheet2')

      expect(xls_w2s.sheets_info).to include(sheet_current: 'Sheet2')
    end

    it 'Change current sheet by sheet index' do
      xls_w2s.preview(current_sheet: 1)

      expect(xls_w2s.sheets_info).to include(sheet_current: 'Sheet2')
    end

    it 'Read xls row 2 ' do
      rows = xls_w2s.preview
      p rows
      expect(rows[1][2..4]).to eq(['Hashimoto', 'Female', 'Great Britain'])
    end

    it 'Read csv row 2 ' do
      rows = csv_eo_wd.preview
      p rows
      expect(rows[2][2..4]).to eq(%w[Gent Männlich Frankreich])
    end
  end
end
# rubocop:enable Metrics/BlockLength
