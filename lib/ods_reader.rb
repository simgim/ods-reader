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
      tv = typed_value cell
      if tv
        cols[index] = tv
        has_value = true
      else
        cols[index] = ''
      end
      colreps = cell.attributes['table:number-columns-repeated']
      if colreps
        colreps.to_i.times do |num|
          cols[index + num] = cols[index]
        end
        index = index + colreps.to_i
      else
        index = index + 1
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

  private
  def typed_value(cell)
    case cell.attributes['office:value-type']
    when nil then nil
    when 'date'
      cell.attributes['office:date-value']
    when 'currency', 'float'
      cell.attributes['office:value']
    else
      if cell.has_elements?
        cell.elements['text:p'].text
      else
        nil
      end
    end
  end
end
