require 'axlsx'

class IndexHelper
  attr_accessor :count

  def initialize
    @count = 0
  end

  def build_index_object(data)
    @count += 1
    i = Index.new
    i.index = @count
    i.question = data.question
    i.sheet_name = data.sheet_name.to_s
    i.filter = data.filters.to_s
    i.error = ""
    i.means = "M" if data.means.present?
    data.header.each do |item|
      i.error = "NOT ESTABLISHED" if item.downcase.include? 'not established'
    end
    data.val.each do |val|
      i.error = "NOT ESTABLISHED" if val['count'][0].downcase.include? 'not established'
    end
    i
  end

  def process_and_write_indexes_to_excel(sheet,indexes,blue_link)
    indexes.each do |index|
      current_row = sheet.add_row([index.question, index.filter, index.index, index.sheet_name, index.means, index.error],:widths=>[40, 40, 4,4,4,10])
      sheet.add_hyperlink location: index.link, ref: "A#{current_row.index+1}", target: :sheet
      sheet["A#{index.index}"].style = blue_link
    end
  end

end