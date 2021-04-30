require_relative 'lib/import_table/version'

Gem::Specification.new do |spec|
  spec.name                  = 'import_table'
  spec.version               = ImportTable::VERSION
  spec.authors               = ['Aleksandr Polyanskiy']
  spec.email                 = ['a.polyanskiy@outlook.com']
  spec.summary               = 'Import table from xlsx and csv.'
  spec.description           = 'Import table from xlsx and csv.'
  spec.homepage              = ''
  spec.license               = 'MIT'
  spec.platform              = Gem::Platform::RUBY
  spec.required_ruby_version = Gem::Requirement.new('>= 2.6.6')
  spec.files                 = Dir.chdir(File.expand_path(__dir__)) do
    'git ls-files -z'.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  end
  spec.require_paths         = ['lib']
  spec.extra_rdoc_files      = %w[README.md LICENSE]

  spec.add_dependency 'roo', '~> 2.8.0'
  spec.add_dependency 'roo-xls', '~> 1.2'
  spec.add_dependency 'ruby-filemagic', '~> 0.7.2'

  spec.add_development_dependency 'rake', '~>13'
  spec.add_development_dependency 'rspec', '~>3.6'
  spec.add_development_dependency 'rubocop', '~> 1.13'
  spec.add_development_dependency 'rubocop-performance', '~> 1.5'
  spec.add_development_dependency 'rubocop-rspec', '~> 2.3'
  spec.add_development_dependency 'simplecov', '~>0.16'
end
