# -*- coding: utf-8 -*-
require 'rake/clean'

EXCEL_FILE  = "sample.xlsm"
MACRO_EXEC_FILE = "test.vbs"

task :default => "test"

desc "Open Excel File"
task :open do
  `cygstart #{EXCEL_FILE}` 
end

desc "Reload All Modules"
task :reload do
  p "to be implemented"
end

desc "Run All Tests"
task :test do
  `cygstart #{MACRO_EXEC_FILE} #{EXCEL_FILE}`
end
