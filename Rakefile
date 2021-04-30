require 'rake'
require 'rake/testtask'
require 'bundler/gem_tasks'
require 'rubocop/rake_task'

RuboCop::RakeTask.new do |t|
  t.requires << 'rubocop-performance'
  t.requires << 'rubocop-rspec'
end

# Rake::TestTask.new(:test) do |t|
#   t.libs << 'test'
#   t.libs << 'lib'
#   t.test_files = FileList['test/**/*_test.rb']
# end

task default: :test
