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
    predefined_styles['normal'] = styles.add_style(:sz => 10, :font_name => 'Tahoma')
    predefined_styles['border'] = styles.add_style(:border => Axlsx::STYLE_THIN_BORDER, :sz => 10, :font_name => 'Tahoma')
    predefined_styles['border_with_center'] = styles.add_style(:border => Axlsx::STYLE_THIN_BORDER,:alignment => { :horizontal => :center }, :sz => 10, :font_name => 'Tahoma')
    predefined_styles['bold'] = styles.add_style(:b => true, :sz => 10, :font_name => 'Tahoma')
    predefined_styles['bold_border'] = styles.add_style(:b => true, :border => Axlsx::STYLE_THIN_BORDER, :sz => 10, :font_name => 'Tahoma',:alignment => { :horizontal => :left })
    predefined_styles['bold_with_center'] = styles.add_style(:b => true, :sz => 10, :font_name => 'Tahoma', :alignment => { :horizontal => :center })
    predefined_styles['bold_border_with_center'] = styles.add_style(:b => true, :border => Axlsx::STYLE_THIN_BORDER , :sz => 10, :font_name => 'Tahoma',:alignment => { :horizontal => :center })
    predefined_styles['red_bold_with_center'] = styles.add_style(:b => true, :fg_color=>"FF0000", :sz => 10, :font_name => 'Tahoma',:alignment => { :horizontal => :center })
    predefined_styles['red_bold_border_with_center'] = styles.add_style(:b => true, :border => Axlsx::STYLE_THIN_BORDER,:fg_color=>"FF0000", :sz => 10, :font_name => 'Tahoma',:alignment => { :horizontal => :center }
    )
    predefined_styles['red_bold_border_with_right'] = styles.add_style(:b => true, :border => Axlsx::STYLE_THIN_BORDER,:fg_color=>"FF0000",:sz => 10, :font_name => 'Tahoma',:alignment => { :horizontal => :right }
    )
    predefined_styles['red_bold_border'] = styles.add_style(:b => true, :border => Axlsx::STYLE_THIN_BORDER, :fg_color=>"FF0000", :sz => 10, :font_name => 'Tahoma'
    )
    predefined_styles['red_bold_border_with_left'] = styles.add_style(:b => true, :border => Axlsx::STYLE_THIN_BORDER,:fg_color=>"FF0000", :sz => 10, :font_name => 'Tahoma',:alignment => { :horizontal => :left }
    )
    predefined_styles['blue_bold_border_with_center'] = styles.add_style(:b => true, :border => Axlsx::STYLE_THIN_BORDER,:fg_color=>"0000FF", :sz => 10, :font_name => 'Tahoma',:alignment => { :horizontal => :center })
    predefined_styles['red'] = styles.add_style(:fg_color=>"FF0000", :sz => 10, :font_name => 'Tahoma')
    predefined_styles['red_bold'] = styles.add_style(:b => true, :fg_color=>"FF0000", :sz => 10, :font_name => 'Tahoma')
    predefined_styles['blue_bold'] = styles.add_style(:b => true, :fg_color=>"0000FF", :sz => 10, :font_name => 'Tahoma')
    predefined_styles['blue_bold_with_center'] = styles.add_style(:b => true, :fg_color=>"0000FF", :sz => 10, :font_name => 'Tahoma', :alignment => { :horizontal => :center })
    predefined_styles
  end

  def add_question(sheet,data,options,predefined_styles)
    p __method__
    sheet.add_row([data.question], style: predefined_styles['blue_bold'],:widths => [:ignore] * data.header.count) if !data.question.nil?
  end

  def add_filters(sheet,data,options,predefined_styles)
    p __method__
    sheet.add_row([data.filters], style: predefined_styles['bold'],:widths => [:ignore] * data.header.count) unless data.filters.nil?
  end

  def add_base(sheet, data, options, predefined_styles)
    p __method__
    sheet.add_row([data.base],style: predefined_styles['normal']) unless data.base.nil?
  end

  def add_weight(sheet, data, options, predefined_styles)
    p __method__
    sheet.add_row([data.weight],style: predefined_styles['normal']) unless data.weight.nil?
  end

  def add_wtd_resp(sheet,data,options,predefined_styles)
    p __method__
    unless data.codelist
      add_row(sheet,data.wtd_resp,options,predefined_styles['blue_bold'],predefined_styles['blue_bold_border_with_center'],data.codelist)
    else
      add_row(sheet,data.wtd_resp,options,predefined_styles['blue_bold_with_center'],predefined_styles['blue_bold_border_with_center'],data.codelist)
    end
  end

  def add_resp(sheet,data,options,predefined_styles)
    p __method__
    unless data.codelist
      add_row(sheet,data.resp,options,predefined_styles['red_bold'],predefined_styles['red_bold_border_with_center'],data.codelist)
    else
      add_row(sheet,data.resp,options,predefined_styles['red_bold_with_center'],predefined_styles['red_bold_border_with_center'],data.codelist)
    end
  end

  def add_header_and_data(sheet,data,options,predefined_styles)
    p __method__
    # Styling Header
    # Styling Data Row
    style_for_header = [predefined_styles['red_bold']]
    style_for_data = [predefined_styles['bold_border']]
    style_for_header_row_2 = [predefined_styles['red_bold']]

    if data.codelist
      if options.present? && options["dual_languages"] == true
        style_for_header = [predefined_styles['red_bold_with_center'],predefined_styles['red_bold_with_center'],predefined_styles['red_bold_with_center']]
        style_for_data = [predefined_styles['bold_border'],predefined_styles['bold_border'],predefined_styles['bold_border']]
        style_for_group_data = [predefined_styles['red_bold_border_with_left'],predefined_styles['red_bold_border'],predefined_styles['red_bold_border']]
        style_for_header_row_2 = [predefined_styles['red_bold'],predefined_styles['red_bold'],predefined_styles['red_bold']]
      else
        style_for_header = [predefined_styles['red_bold_with_center'],predefined_styles['red_bold_with_center']]
        style_for_data = [predefined_styles['bold_border'],predefined_styles['bold_border']]
        style_for_group_data = [predefined_styles['red_bold_border_with_left'],predefined_styles['red_bold_border']]
        style_for_header_row_2 = [predefined_styles['red_bold'],predefined_styles['red_bold']]
      end
    end

    old = []

    if options['export_data_type'] == :both
      if data.codelist
        if options.present? && options["dual_languages"] == true
          ((data.header.count - 3)*2).times { style_for_header << predefined_styles['red_bold_border_with_center'] }
          ((data.header.count - 3)*2).times { style_for_data << predefined_styles['border_with_center'] }
          ((data.header.count - 3)*2).times { style_for_group_data << predefined_styles['red_bold_border_with_center'] }
          ((data.header.count - 3)*2).times { style_for_header_row_2 << predefined_styles['red_bold_border_with_center'] }
        else
          ((data.header.count - 2)*2).times { style_for_header << predefined_styles['red_bold_border_with_center'] }
          ((data.header.count - 2)*2).times { style_for_data << predefined_styles['border_with_center'] }
          ((data.header.count - 2)*2).times { style_for_group_data << predefined_styles['red_bold_border_with_center'] }
          ((data.header.count - 2)*2).times { style_for_header_row_2 << predefined_styles['red_bold_border_with_center'] }
        end
      else
        ((data.header.count - 1)*2).times { style_for_header << predefined_styles['red_bold_border_with_right'] }
        ((data.header.count - 1)*2).times { style_for_data << predefined_styles['border_with_center'] }
        ((data.header.count - 1)*2).times { style_for_header_row_2 << predefined_styles['red_bold_border_with_center']}
      end
      old = data.header.dup
      data.header = fill_blanks(data.header,data.codelist)
    else
      if data.codelist
        if options.present? && options["dual_languages"] == true
          (data.header.count - 3).times { style_for_header << predefined_styles['red_bold_with_center'] }
          (data.header.count - 3).times { style_for_data << predefined_styles['border'] }
          (data.header.count - 3).times { style_for_group_data << predefined_styles['red_bold_border_with_right'] }
          (data.header.count - 3).times { style_for_header_row_2 << predefined_styles['red_bold_border_with_center'] }
        else
          (data.header.count - 2).times { style_for_header << predefined_styles['red_bold_with_center'] }
          (data.header.count - 2).times { style_for_data << predefined_styles['border'] }
          (data.header.count - 2).times { style_for_group_data << predefined_styles['red_bold_border_with_right'] }
          (data.header.count - 2).times { style_for_header_row_2 << predefined_styles['red_bold_border_with_center'] }
        end
      else
        (data.header.count - 1).times { style_for_header << predefined_styles['red_bold_with_center'] }
        (data.header.count - 1).times { style_for_data << predefined_styles['border'] }
        (data.header.count - 1).times { style_for_header_row_2 << predefined_styles['red_bold_border_with_center'] }
      end
    end

    width = [:ignore] * data.header.count
    if data.codelist
      if options.present? && options["dual_languages"] == true
        width = [:ignore, 40, 40]
        width << [:ignore] * (data.header.count - 3)
      else
        width = [:ignore, 40]
        width << [:ignore] * (data.header.count - 2)
      end
      width = width.flatten
    end
    start_row = sheet.add_row(data.header, style: style_for_header,:widths => width)

    if options['export_data_type'] == :both
      merge_cells(sheet,start_row,old,predefined_styles['red_bold_border_with_center'],data.codelist)
      new_row = []
      new_row[0] = ''
      i = 1
      j = 1

      if data.codelist
        if options.present? && options["dual_languages"] == true
          i = 3
          j = 3
        else
          i = 2
          j = 2
        end
      end

      if options['export_first'] == :percent
        (old.count - j).times {
          new_row[i] = 'percent'
          new_row[i+1] = 'count'
          i += 2
        }
      else
        (old.count - j).times {
          new_row[i] = 'count'
          new_row[i+1] = 'percent'
          i += 2
        }
      end

      sheet.add_row(new_row, style: style_for_header_row_2,:widths => [:ignore] * data.header.count)
    end

    ap 'Write Value'
    data.val.each do |value|
      if value == nil
        next
      end
      value_arr = []
      value_arr = value['count'] if options['export_data_type'] == :count
      value_arr = value['percent'] if options['export_data_type'] == :percent
      if options['export_data_type'] == :both
        if options['export_first'] == :percent
          value_arr = merge_arr(value['count'],value['percent'],true,data.codelist)
        else
          value_arr = merge_arr(value['count'],value['percent'],false,data.codelist)
        end
      end
      if value['bold'].present? && value['bold']
        sheet.add_row(value_arr, style: style_for_group_data,:widths => [:ignore] * data.header.count)
      else
        sheet.add_row(value_arr, style: style_for_data,:widths => [:ignore] * data.header.count)
      end
    end

    last_cell = sheet.rows.last.cells.last

    if options['export_data_type'] != :both
      sheet.auto_filter = "#{Axlsx::cell_r(0,start_row.index)}:#{Axlsx::cell_r(last_cell.index,last_cell.row.index)}"
    end

  end

  def add_totals(sheet,data,options,predefined_styles)
    p __method__
    styling = [predefined_styles['bold']]

    if data.codelist
      if options.present? && options["dual_languages"] == true
        styling = [predefined_styles['bold_with_center'],predefined_styles['bold_with_center'],predefined_styles['bold_with_center']]
      else
        styling = [predefined_styles['bold_with_center'],predefined_styles['bold_with_center']]
      end
    end

    total_arr = data.totals_count if options['export_data_type'] == :count
    total_arr = data.totals_percent if options['export_data_type'] == :percent
    if options['export_data_type'] == :both
      if options['export_first'] == :percent
        total_arr = merge_arr(data.totals_count,data.totals_percent,true,data.codelist)
      else
        total_arr = merge_arr(data.totals_count,data.totals_percent,false,data.codelist)
      end
    end
    s = predefined_styles['bold']

    if options['export_data_type'] == :both
      s = predefined_styles['bold_border_with_center']
    end

    i = 1
    if data.codelist
      if options.present? && options["dual_languages"] == true
        i = 3
      else
        i = 2
      end
    end

    (total_arr.count-i).times {
      styling << s
    }

    row = sheet.add_row(total_arr,style: styling)
    if options['export_data_type'] == :both
      if data.codelist
        if options.present? && options["dual_languages"] == true
          range = Axlsx::cell_r(0,row.index) + ":" + Axlsx::cell_r(2, row.index)
        else
          range = Axlsx::cell_r(0,row.index) + ":" + Axlsx::cell_r(1, row.index)
        end

        sheet.merge_cells range
      end
    end
  end

  def add_std_deviation(sheet,data,options,predefined_styles)
    p __method__
    unless data.codelist
      add_row(sheet,data.std_deviation,options,predefined_styles['bold'],predefined_styles['bold_border_with_center'],data.codelist)
    else
      add_row(sheet,data.std_deviation,options,predefined_styles['bold_with_center'],predefined_styles['bold_border_with_center'],data.codelist)
    end
  end

  def add_means(sheet,data,options,predefined_styles)
    p __method__
    unless data.codelist
      add_row(sheet,data.means,options,predefined_styles['bold'],predefined_styles['bold_border_with_center'],data.codelist)
    else
      add_row(sheet,data.means,options,predefined_styles['bold_with_center'],predefined_styles['bold_border_with_center'],data.codelist)
    end
  end

  def add_mode(sheet,data,options,predefined_styles)
    p __method__
    unless data.codelist
      add_row(sheet,data.mode,options,predefined_styles['bold'],predefined_styles['bold_border_with_center'],data.codelist)
    else
      add_row(sheet,data.mode,options,predefined_styles['bold_with_center'],predefined_styles['bold_border_with_center'],data.codelist)
    end
  end

  def add_medians(sheet,data,options,predefined_styles)
    p __method__
    unless data.codelist
      add_row(sheet,data.medians,options,predefined_styles['bold'],predefined_styles['bold_border_with_center'],data.codelist)
    else
      add_row(sheet,data.medians,options,predefined_styles['bold_with_center'],predefined_styles['bold_border_with_center'],data.codelist)
    end
  end

  def merge_arr(count_arr,percent_arr,percent_first=false,codelist=false)
    both_arr = [count_arr[0]]
    index = 1
    if codelist
      both_arr << count_arr[1]
      index = 2
    end
    count_arr[index..-1].each.with_index(index) do |v,i|
      if percent_first == true
        both_arr << percent_arr[i] unless percent_arr.nil?
        both_arr << count_arr[i] unless count_arr.nil?

      else
        both_arr << count_arr[i] unless count_arr.nil?
        both_arr << percent_arr[i] unless percent_arr.nil?

      end
      if percent_arr.nil?
        both_arr << 0
      end
    end
    both_arr
  end

  def merge_cells(sheet,row,data_arr,style,codelist=false,options=nil)
    p __method__
    if codelist
      range = Axlsx::cell_r(0,row.index) + ":" + Axlsx::cell_r(1, row.index)
      sheet.merge_cells range
    end
    codelist ? index = 2 : index = 1
    index = 3 if options.present? && options["dual_languages"] == true
    data_arr[index..-1].each do
      range = Axlsx::cell_r(index,row.index) + ":" + Axlsx::cell_r(index+1, row.index)
      sheet.merge_cells range
      # p range
      sheet[Axlsx::cell_r(index,row.index) + ":" + Axlsx::cell_r(index+1, row.index)].each { |c|
        c.style = style
      }
      index = index + 2
    end
  end

  def fill_blanks(data_arr,codelist=false,options=nil)
    p __method__
    codelist ? index = 2 : index = 1
    index = 3 if options.present? && options["dual_languages"] == true
    data_arr[index..-1].each do
      data_arr = data_arr.insert(index+1,'')
      index = index + 2
    end
    data_arr
  end

  def add_row(sheet,data,options,predefined_styles_1,predefined_styles_2,codelist=false)
    p __method__
    unless data.nil?
      old_data = data.dup
      if options['export_data_type'] == :both
        data = fill_blanks(data,codelist)
      end
      row = sheet.add_row(data, style: predefined_styles_1)
      if options['export_data_type'] == :both
        merge_cells(sheet,row,old_data,predefined_styles_2,codelist)
      end
    end
  end

end