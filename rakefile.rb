# -*- coding: utf-8 -*-
require 'rake/clean'

FILE_PATH  = "./test.xlsm"

task :default => "open"

task :open do
  `cygstart  #{FILE_PATH}` 
end

task :test do
  `cygstart test.vbs` 
end
