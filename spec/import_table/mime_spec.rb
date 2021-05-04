require_relative '../spec_helper'

TEST_MIMETYPE = {
  'Define file type - > .csv'  => { expect: :csv, file: 'file_example_CSV_10_bar.csv' },
  'Define file type - > .ods'  => { expect: :ods, file: 'file_example_ODS_10.ods' },
  'Define file type - > .tsv'  => { expect: :csv, file: 'file_example_TSV_10.tsv' },
  'Define file type - > .xls'  => { expect: :xls, file: 'file_example_XLS_10.xls' },
  'Define file type - > .xlsx' => { expect: :xlsx, file: 'file_example_XLSX_10.xlsx' }
}.freeze

TEST_DELIMITERS_SHORT = {
  'Define delimiter -> comma'     => { expect: ',', file: 'file_example_CSV_10_comma.csv' },
  'Define delimiter -> semicolon' => { expect: ';', file: 'file_example_CSV_10_semicolon.csv' }
}.freeze

# rubocop:disable Metrics/BlockLength
describe ImportTable::Mime do
  let(:csv_de) { described_class.new(get_file('file_example_CSV_10_comma_de.csv')) }

  describe '.new' do
    it 'Creates an instance' do
      expect(csv_de).to be_a(described_class)
    end

    it 'Creates an instance from StringIO' do
      param     = TEST_MIMETYPE['Define file type - > .tsv']
      test_mime = described_class.new(StringIO.new(File.open(get_file(param[:file])).read))

      expect(test_mime).to be_a(described_class)
    end

    it 'Cannot open file -> exception' do
      expect { described_class.new(get_file('NoFile.xlsx')) }.to raise_error(ImportTable::FileCannotOpen)
    end
  end

  describe '.type?' do
    TEST_MIMETYPE.each do |name, param|
      it name do
        expect(described_class.new(get_file(param[:file])).type?).to eq(param[:expect])
      end
    end

    it 'Unsupported file type' do
      file = get_file('file_example_RTF_100kB.rtf')

      expect { described_class.new(file).type? }.to raise_error(ImportTable::UnsupportedType)
    end
  end

  describe '.delimiter?' do
    TEST_DELIMITERS_SHORT.each do |name, param|
      it name do
        expect(described_class.new(get_file(param[:file])).delimiter?).to eq(param[:expect])
      end
    end
  end

  describe '.encoding?' do
    it 'Encoding -> UTF-8' do
      expect(csv_de.encoding?).to eq('utf-8')
    end
  end
end
# rubocop:enable Metrics/BlockLength
