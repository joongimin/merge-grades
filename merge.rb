#!/usr/bin/env ruby

require 'roo'
require 'csv'

dirs = Dir['data/*']

csv_str = CSV.generate do |csv|
  csv << dirs.map {|dir| dir[5..-1]}
  cols = dirs.map do |dir|
    file = "#{dir}/Final Scoring Sheet.xlsx"
    xlsx = Roo::Spreadsheet.open(file, extension: :xlsx)
    xlsx.default_sheet = xlsx.sheets.first
    (5..59).map {|row| xlsx.cell(row, 3)}
  end
  55.times do |i|
    csv << cols.map {|c| c[i]}
  end
end

File.open('output.csv', 'w') {|f| f.write csv_str}
