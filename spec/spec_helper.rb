require 'simplecov'

SimpleCov.start { add_filter 'spec' }

# SimpleCov.minimum_coverage 99

require 'import_table'

DATA_PATH = "#{File.dirname(__FILE__)}/data".freeze

def get_file(file_name)
  File.join(DATA_PATH, file_name)
end

def get_string_io(file_name)
  StringIO.new(File.open(get_file(file_name)).read)
end
