module ImportTable
  # A base error class for ImportTable
  class Error < StandardError; end

  # Raised when ImportTable::Mime cannot open file
  class FileCannotOpen < Error; end

  class UnsupportedFileType < Error; end

  class MissingRequiredOption < Error; end

  class SheetNotFound < Error; end
end
