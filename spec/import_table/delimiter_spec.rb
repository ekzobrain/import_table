require_relative '../spec_helper'

TEST_DELIMITERS = {
  'Define delimiter -> bar'       => { expect: '|', file: 'file_example_CSV_10_bar.csv' },
  'Define delimiter -> colon'     => { expect: ':', file: 'file_example_CSV_10_colon.csv' },
  'Define delimiter -> semicolon' => { expect: ';', file: 'file_example_CSV_10_semicolon.csv' },
  'Define delimiter -> tab'       => { expect: "\t", file: 'file_example_CSV_10_tab.csv' },
  'Without delimiter'             => { expect: nil, file: 'file_example_CSV_10_without_delimiter.csv' }
}.freeze

describe ImportTable::Delimiter do
  describe '.type' do
    TEST_DELIMITERS.each do |name, param|
      it name do
        expect(described_class.type(get_file(param[:file]))).to eq(param[:expect])
      end
    end

    it 'Define delimiter from StringIO' do
      param          = TEST_DELIMITERS['Define delimiter -> semicolon']
      test_delimiter = described_class.type(StringIO.new(File.open(get_file(param[:file])).read))

      expect(test_delimiter).to eq(param[:expect])
    end
  end
end
