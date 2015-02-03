require 'axlsx'
# https://github.com/randym/axlsx/blob/master/examples/example.rb
class WriteExcelService < BaseService

  def export_data_to_excel(wb,data,options={},sheet)
    # Setup Styles
    ap data.sheet_name
    predefined_styles = set_predefined_styles(wb)
    sorted_order = options['orders'].sort_by{ |k,v| v }

    # Started to Write Data
    start_row = add_blank_row(sheet)
    sorted_order.each do |key,value|
      if value > 0
        eval("add_#{key}(sheet,data,options,predefined_styles)")
      end
    end
    3.times { add_blank_row(sheet) }
    # sheet.col_style 0, predefined_styles['alignment_left']
    start_row.index + 2
  end

  protected

  def add_blank_row(sheet)
    p __method__
    sheet.add_row()
  end

  def set_predefined_styles(wb)
    p __method__
    styles = wb.styles
    predefined_styles = Hash.new
    predefined_styles['border'] = styles.add_style(:border => Axlsx::STYLE_THIN_BORDER)
    predefined_styles['border_with_center'] = styles.add_style(:border => Axlsx::STYLE_THIN_BORDER,:alignment => { :horizontal => :center })
    predefined_styles['bold'] = styles.add_style(:b => true)
    predefined_styles['bold_border'] = styles.add_style(:b => true, :border => Axlsx::STYLE_THIN_BORDER,:alignment => { :horizontal => :left })
    predefined_styles['bold_border_with_center'] = styles.add_style(:b => true, :border => Axlsx::STYLE_THIN_BORDER,:alignment => { :horizontal => :center })
    predefined_styles['red_bold_with_center'] = styles.add_style(:b => true, :fg_color=>"FF0000",:alignment => { :horizontal => :center })

    predefined_styles['red_bold_boder_with_center'] = styles.add_style(:b => true, :border => Axlsx::STYLE_THIN_BORDER,:fg_color=>"FF0000",:alignment => { :horizontal => :center })

    predefined_styles['blue_bold_boder_with_center'] = styles.add_style(:b => true, :border => Axlsx::STYLE_THIN_BORDER,:fg_color=>"0000FF",:alignment => { :horizontal => :center })

    predefined_styles['red_bold'] = styles.add_style(:b => true, :fg_color=>"FF0000")
    predefined_styles['red'] = styles.add_style(:fg_color=>"FF0000")
    predefined_styles['blue_bold'] = styles.add_style(:b => true, :fg_color=>"0000FF")
    # predefined_styles['alignment_left'] = styles.add_style(:alignment => { :horizontal => :left })
    predefined_styles
  end

  def add_question(sheet,data,options,predefined_styles)
    p __method__
    sheet.add_row([data.question], widths: [:ignore] * data.header.count,
                    style: predefined_styles['bold']) if !data.question.nil?
  end

  def add_filters(sheet,data,options,predefined_styles)
    p __method__
    sheet.add_row([data.filters], style: predefined_styles['blue_bold']) unless data.filters.nil?
  end

  def add_base(sheet,data,options,predefined_styles)
    p __method__
    sheet.add_row([data.base]) unless data.base.nil?
  end

  def add_wtd_resp(sheet,data,options,predefined_styles)
    p __method__
    add_row(sheet,data.wtd_resp,options,predefined_styles['blue_bold'],predefined_styles['blue_bold_boder_with_center'])
  end

  def add_resp(sheet,data,options,predefined_styles)
    p __method__
    add_row(sheet,data.resp,options,predefined_styles['red_bold'],predefined_styles['red_bold_boder_with_center'])
  end

  def add_header_and_data(sheet,data,options,predefined_styles)
    p __method__
    # Styling Header
    # Styling Data Row

    style_for_header = [predefined_styles['red_bold']]
    style_for_data = [predefined_styles['bold_border']]
    old = []

    if options['export_data_type'] == :both
      ((data.header.count - 1)*2).times { style_for_header << predefined_styles['red_bold_boder_with_center'] }
      ((data.header.count - 1)*2).times { style_for_data << predefined_styles['border_with_center'] }
      old = data.header.dup
      data.header = fill_blanks(data.header)
    else
      (data.header.count - 1).times { style_for_header << predefined_styles['red_bold_with_center'] }
      (data.header.count - 1).times { style_for_data << predefined_styles['border'] }
    end

    start_row = sheet.add_row(data.header, style: style_for_header,:widths => [:ignore] * data.header.count)

    if options['export_data_type'] == :both
      merge_cells(sheet,start_row,old,predefined_styles['red_bold_boder_with_center'])
      new_row = []
      new_row[0] = ''
      i = 1
      (old.count - 1).times {
        new_row[i] = 'count'
        new_row[i+1] = 'percent'
        i += 2
      }
      sheet.add_row(new_row, style: style_for_header,:widths => [:ignore] * data.header.count)
    end

    data.val.each do |value|
      value_arr = []
      value_arr = value['count'] if options['export_data_type'] == :count
      value_arr = value['percent'] if options['export_data_type'] == :percent
      value_arr = merge_arr(value['count'],value['percent']) if options['export_data_type'] == :both
      sheet.add_row(value_arr, style: style_for_data)
    end

    last_cell = sheet.rows.last.cells.last

    if options['export_data_type'] != :both
      sheet.auto_filter = "#{Axlsx::cell_r(0,start_row.index)}:#{Axlsx::cell_r(last_cell.index,last_cell.row.index)}"
    end
  end

  def add_totals(sheet,data,options,predefined_styles)
    p __method__
    styling = [predefined_styles['bold']]

    total_arr = data.totals_count if options['export_data_type'] == :count
    total_arr = data.totals_percent if options['export_data_type'] == :percent
    total_arr = merge_arr(data.totals_count,data.totals_percent) if options['export_data_type'] == :both
    # binding.pry
    s = predefined_styles['bold']

    if options['export_data_type'] == :both
      s = predefined_styles['bold_border_with_center']
    end

    (total_arr.count-1).times {
      styling << s
    }

    sheet.add_row(total_arr,style: styling)
  end

  def add_std_deviation(sheet,data,options,predefined_styles)
    p __method__
    add_row(sheet,data.std_deviation,options,predefined_styles['bold'],predefined_styles['bold_border_with_center'])
  end

  def add_means(sheet,data,options,predefined_styles)
    p __method__
    add_row(sheet,data.means,options,predefined_styles['bold'],predefined_styles['bold_border_with_center'])
  end

  def add_mode(sheet,data,options,predefined_styles)
    p __method__
    add_row(sheet,data.mode,options,predefined_styles['bold'],predefined_styles['bold_border_with_center'])
  end

  def add_medians(sheet,data,options,predefined_styles)
    p __method__
    add_row(sheet,data.medians,options,predefined_styles['bold'],predefined_styles['bold_border_with_center'])
  end

  def merge_arr(count_arr,percent_arr)
    both_arr = [count_arr[0]]
    count_arr[1..-1].each.with_index(1) do |v,i|
      both_arr << count_arr[i] unless count_arr.nil?
      both_arr << percent_arr[i] unless percent_arr.nil?
      if percent_arr.nil?
        both_arr << 0
      end
    end
    both_arr
  end

  def merge_cells(sheet,row,data_arr,style)
    p __method__
    index = 1
    data_arr[1..-1].each do
      range = Axlsx::cell_r(index,row.index) + ":" + Axlsx::cell_r(index+1, row.index)
      sheet.merge_cells range
      # p range
      sheet[Axlsx::cell_r(index,row.index) + ":" + Axlsx::cell_r(index+1, row.index)].each { |c|
        c.style = style
      }
      index = index + 2
    end
  end

  def fill_blanks(data_arr)
    p __method__
    index = 1
    data_arr[1..-1].each do
      data_arr = data_arr.insert(index+1,'')
      index = index + 2
    end
    data_arr
  end

  def add_row(sheet,data,options,predefined_styles_1,predefined_styles_2)
    p __method__
    unless data.nil?
      old_data = data.dup
      if options['export_data_type'] == :both
        data = fill_blanks(data)
      end
      row = sheet.add_row(data, style: predefined_styles_1)
      if options['export_data_type'] == :both
        merge_cells(sheet,row,old_data,predefined_styles_2)
      end
    end
  end

end