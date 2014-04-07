# -*- coding: utf-8 -*-
require 'rake/clean'
require 'win32ole'

EXCEL_FILE  = "sample.xlsm"
MACRO_EXEC_FILE = "test.vbs"
DEBUG_SHOW = true  

task :default => "test"

desc "Open or Connect Excel File"
task :open do
  @xl = openExcel(EXCEL_FILE)
end

desc "Reload All Modules"
task :reload => :open do
  @xl.run("ThisWorkBook.reloadModule")
end

desc "Run All Tests"
task :test do
  `cygstart #{MACRO_EXEC_FILE} #{EXCEL_FILE}`
end

# refered from 
# http://osdir.com/ml/lang.ruby.japanese/2005-11/msg00180.html
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end

def openExcel(filename)
  filename = getAbsolutePath(filename)
  xl = nil
  begin
    xl = WIN32OLE::connect("Excel.Application")
  rescue WIN32OLERuntimeError
    xl = WIN32OLE.new("Excel.Application")
  end
  xl.Workbooks.each do |sheet|
    if sheet.FullName == filename
      sheet.Activate
    end
  end

  unless xl.ActiveWorkbook && xl.ActiveWorkbook.FullName == filename
    xl.Workbooks.Open(filename)
  end
  xl.Visible = true
  return xl
end
