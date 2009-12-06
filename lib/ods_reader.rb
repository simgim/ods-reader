require 'rexml/document'
require 'zip/zipfilesystem'

class ODSReader

  include REXML

  attr_accessor :skip_empty_row

  def self.open(filename, sheet_name, &block)
    reader = new
    reader.skip_empty_row = false
    reader.process_book(filename, sheet_name, block)
  end

  def process_book(filename, sheet_name, block)
    docbytes = nil
    Zip::ZipFile.open(filename) do |zipfile|
      docbytes = zipfile.file.read('content.xml')
    end

    doc = Document.new(docbytes)
    ws = doc.elements["//table:table[@table:name='#{sheet_name}']"]
    process_sheet(ws, block)
  end

  def process_sheet(sheet, block)
    sheet.elements.each('table:table-row') do |row|
      rowreps = row.attributes['table:number-rows-repeated'] || '1'
      rowreps = rowreps.to_i
      process_row(rowreps, row, block)
    end
  end

  def process_row(rowreps, row, block)
    cols = []
    index = 0
    has_value = false
    row.elements.each('table:table-cell') do |cell|
      colreps = cell.attributes['table:number-columns-repeated']
      unless colreps
        unless cell.has_elements?
          colreps = '1'
        end
      end
      if colreps
        colreps.to_i.times do |num|
          cols[index] = ''
          index = index + 1
        end
      else
        cols[index] = cell.elements['text:p'].text
        index = index + 1
        has_value = true
      end
    end
    rowreps.times do |num|
      if has_value
        block.call(cols)
      elsif !skip_empty_row
        block.call(cols)
      end
    end
  end
end
