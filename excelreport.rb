require 'caxlsx'
require 'csv'
require 'yaml'
require 'digest'
require 'ipaddr'

class ExcelReport
  def initialize(conf_filename)
    @time_stamp  = Time.now
    @pretty_date = "%a, %b %-d, %Y %l:%M %p"
    @moment      = "%Y%m%d-%H%M%S"
    @report_conf = YAML.load_file(conf_filename)
    @report      = Axlsx::Package.new
    @workbook    = @report.workbook
    @report_conf["sheets"].each_with_index { |sheet, idx| @report_conf["sheets"][idx]["filename"] = Dir.glob("#{sheet["data_directory"]}/#{sheet["base_filename"]}*")[-1] }
    @report_conf["sheets"].each_with_index { |sheet, idx| @report_conf["sheets"][idx]["csv_table"] = CSV.read(sheet["filename"], headers: true) }
    @report_conf["sheets"].each { |sheet| add_worksheet(sheet) }
    @report_conf["sheets"].each_with_index { |sheet, idx| conf_sheet_view(sheet, idx) }
  end

  def add_worksheet(worksheet)
    sheet = @workbook.add_worksheet(name: worksheet["worksheet_name"])
    default_font = { font_name: 'Calibri', sz: 12 }
    default_style = sheet.styles.add_style(default_font)
    worksheet["total_row_styles"] = worksheet["column_styles"].dup
    worksheet["column_styles"].each do |column, style_hash|
      if style_hash && style_hash["default_style"]
        worksheet["column_styles"][column] = sheet.styles.add_style(default_font.merge(style_hash["style"]))
        worksheet["total_row_styles"][column] = default_font.merge(style_hash["style"])
      elsif style_hash
        worksheet["column_styles"][column] = sheet.styles.add_style(style_hash["style"])
      else
        worksheet["column_styles"][column] = default_style
        worksheet["total_row_styles"][column] = default_font
      end
    end
    total_row_styles = conf_total_row_styles(worksheet["table_styles"][:name])
    worksheet["total_row_styles"].each do |column, style_hash|
      if worksheet["total_row_styles"].keys[0] == column
        worksheet["total_row_styles"][column] = sheet.styles.add_style(total_row_styles[:left].merge(style_hash))
      elsif worksheet["total_row_styles"].keys[-1] == column
        worksheet["total_row_styles"][column] = sheet.styles.add_style(total_row_styles[:right].merge(style_hash))
      else
        worksheet["total_row_styles"][column] = sheet.styles.add_style(total_row_styles[:middle].merge(style_hash))
      end
    end
    col_index = 'A'
    row_index = 1
    header_styles = conf_header_styles(worksheet["table_styles"][:name])
    if worksheet["header"]
      header_row = 0
      worksheet["header"].each do |label, field|
        if field["format"] && field["format"]["default_style"]
          label_style       = sheet.styles.add_style(default_font.merge(header_styles[:labels]))
          value_odd_style   = sheet.styles.add_style(default_font.merge(header_styles[:value_odd].merge(field["format"]["style"])))
          value_even_style  = sheet.styles.add_style(default_font.merge(header_styles[:value_even].merge(field["format"]["style"])))
        elsif field["format"]
          label_style       = sheet.styles.add_style(header_styles[:labels])
          value_odd_style   = sheet.styles.add_style(header_styles[:value_odd].merge(field["format"]["style"]))
          value_even_style  = sheet.styles.add_style(header_styles[:value_even].merge(field["format"]["style"]))
        else
          label_style       = sheet.styles.add_style(default_font.merge(header_styles[:labels]))
          value_odd_style   = sheet.styles.add_style(default_font.merge(header_styles[:value_odd]))
          value_even_style  = sheet.styles.add_style(default_font.merge(header_styles[:value_even]))
        end
        value_style = header_row % 2 == 0 ? value_even_style : value_odd_style
        unless field["type"] == "blank"
          if field["type"] == "time_stamp"
            sheet.add_row [label, @time_stamp.strftime(@pretty_date)], style: [label_style, value_style]
          else
            sheet.add_row [label, field["value"]], style: [label_style, value_style]
          end
        else
          sheet.add_row [], style: default_style
        end
        row_index += 1
        header_row += 1
      end
    else
      default_header_style = sheet.styles.add_style(default_font.merge(header_styles[:labels]))
      sheet.add_row [@time_stamp.strftime(@pretty_date)], style: default_header_style
      sheet.add_row [], style: default_style
      row_index += 2
    end
    if worksheet["calculated_sheet"]
      unless worksheet["calculated_table"].select { |field, value| value["field_type"] == "collected" }.empty?
        table_hash = Hash.new do |hash, key|
          hash[key] = if worksheet["calculated_table"][key]["field_type"] == "collected"
            eval worksheet["calculated_table"][key]["collection_statement"]
          else
            worksheet["calculated_table"][key]["formula"]
          end
        end
        worksheet["calculated_table"].each { |header, field| table_hash[header] }
        table_build = CSV::Table.new([])
        collected_key = worksheet["calculated_table"].select { |header, field| field["field_type"] == "collected" }.keys[0]
        formulas = table_hash.collect { |header, value| value.class == Array ? nil : value }
        collected_index = formulas.index(nil)
        table_hash[collected_key].each do |value|
          formulas[collected_index] = value
          row_vals = formulas
          table_build << CSV::Row.new(table_hash.keys, row_vals)
        end
        worksheet["csv_table"] = table_build
        sheet.add_row worksheet["csv_table"].headers, style: default_style
        worksheet["csv_table"].each { |row| sheet.add_row row.values_at, style: worksheet["column_styles"].values }
      else
        sheet.add_row ["CALCULATED TABLE CONFIGURATION ERROR: NO DESIGNATED \"COLLECTED\" COLUMN"], style: default_style
      end
    else
      sheet.add_row worksheet["csv_table"].headers, style: default_style
      worksheet["csv_table"].each { |row| sheet.add_row row.values_at, style: worksheet["column_styles"].values }
    end
    sheet.add_row worksheet["total_row"].values, style: worksheet["total_row_styles"].values if worksheet["total_row"]
    sheet.add_table "#{range(worksheet["csv_table"], "#{col_index}#{row_index}")}", name: "#{worksheet["worksheet_name"].downcase.split(" ").join("_")}", style_info: worksheet["table_styles"]
    if worksheet["column_widths"]
      column_widths_to_s = -> (column_width_array) do
        string_value = String.new
        column_width_array.each_with_index do |elem, idx|
          if idx == column_width_array.size - 1
            string_value << (elem.nil? ? "nil" : elem.to_s)
          else
            string_value << (elem.nil? ? "nil, " : "#{elem}, ")
          end
        end
        string_value
      end
      set_widths_string = "sheet.column_widths " + column_widths_to_s.call(worksheet["column_widths"].values)
      eval set_widths_string
    end
    if worksheet["freeze_panes"]
      top_left_cell = if worksheet["freeze_panes"] == "table"
        row_index += 1 if worksheet["total_row"]
        next_col(next_row(range(worksheet["csv_table"], "#{col_index}#{row_index}").split(":")[-1]))
      else
        worksheet["freeze_panes"]
      end
      coordinates = coords(top_left_cell)
      sheet.sheet_view.pane do |pane|
        pane.state = :frozen
        pane.y_split = coordinates[:y_split]
        pane.x_split = coordinates[:x_split]
      end
    end
    sheet.sheet_view.add_selection(:top_left, { active_cell: worksheet["active_cell"], sqref: worksheet["active_cell"] }) if worksheet["active_cell"]
  end

  def range(csv_table, start_cell)
    start_col = start_cell.match(/(?<alpha>[A-Z])/)[:alpha]
    start_row = start_cell.match(/(?<num>[0-9]+)/)[:num].to_i
    stop_col = start_col.dup
    (csv_table.headers.size - 1).times { stop_col.succ! }
    stop_row = start_row
    csv_table.size.times { stop_row += 1 }
    start_col + start_row.to_s + ":" + stop_col + stop_row.to_s
  end
  
  def next_col(cell_ref)
    col = cell_ref.match(/(?<alpha>[A-Z])/)[:alpha]
    row = cell_ref.match(/(?<num>[0-9]+)/)[:num].to_i
    col = col.succ
    col + row.to_s
  end

  def next_row(cell_ref)
    col = cell_ref.match(/(?<alpha>[A-Z])/)[:alpha]
    row = cell_ref.match(/(?<num>[0-9]+)/)[:num].to_i
    row += 1
    col + row.to_s
  end

  def coords(cell_ref)
    col = cell_ref.match(/(?<alpha>[A-Z])/)[:alpha]
    row = cell_ref.match(/(?<num>[0-9]+)/)[:num].to_i
    coords = { y_split: 0, x_split: 0 }
    start_col = 'A'
    start_row = 1
    coords[:y_split] = row - start_row
    until start_col == col
      start_col = start_col.succ
      coords[:x_split] += 1
    end
    coords
  end

  def conf_sheet_view(worksheet, index)
    if not worksheet["sheet_view"].nil?
      worksheet["sheet_view"].each do |key, value|
        eval "@workbook.worksheets[#{index}].sheet_view.#{key} = #{value}"
      end
    end
  end

  def conf_header_styles(table_style_string = "default")
    style_hash = Hash.new do |hash, key|
      hash[key] = { border: { style: :thin, edges: [:left, :top, :right, :bottom] } }
    end
    [:labels, :value_odd, :value_even].each { |header_part| style_hash[header_part] }
    stylize = -> (style_spec) do
      style_hash.each { |header_part, style| style[:border] = style[:border].merge(style_spec[:border]) }
      style_hash[:labels]     = style_hash[:labels].merge(style_spec[:labels])
      style_hash[:value_odd]  = style_hash[:value_odd].merge(style_spec[:value_odd])
      style_hash[:value_even] = style_hash[:value_even].merge(style_spec[:value_even])
    end
    style_spec = case table_style_string
    when "TableStyleMedium1"
      {
        border:     { color: '000000' },
        labels:     { bg_color: '000000', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'D9D9D9', fg_color: '000000' },
        value_even: { fg_color: '000000' }
      }
    when "TableStyleMedium2"
      {
        border:     { color: '8EA9DB' },
        labels:     { bg_color: '4472C4', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'D9E1F2', fg_color: '000000' },
        value_even: { fg_color: '000000' }
      }
    when "TableStyleMedium3"
      {
        border:     { color: 'F4B084' },
        labels:     { bg_color: 'ED7D31', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'FCE4D6', fg_color: '000000' },
        value_even: { fg_color: '000000' }
      }
    when "TableStyleMedium4"
      {
        border:     { color: 'C9C9C9' },
        labels:     { bg_color: 'A5A5A5', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'EDEDED', fg_color: '000000' },
        value_even: { fg_color: '000000' }
      }
    when "TableStyleMedium5"
      {
        border:     { color: 'FFD966' },
        labels:     { bg_color: 'FFC000', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'FFF2CC', fg_color: '000000' },
        value_even: { fg_color: '000000' }
      }
    when "TableStyleMedium6"
      {
        border:     { color: '9BC2E6' },
        labels:     { bg_color: '5B9BD5', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'DDEBF7', fg_color: '000000' },
        value_even: { fg_color: '000000' }
      }
    when "TableStyleMedium7"
      {
        border:     { color: 'A9D08E' },
        labels:     { bg_color: '70AD47', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'E2EFDA', fg_color: '000000' },
        value_even: { fg_color: '000000' }
      }
    when "TableStyleMedium8"
      {
        border:     { color: 'FFFFFF' },
        labels:     { bg_color: '000000', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'A6A6A6', fg_color: '000000' },
        value_even: { bg_color: 'D9D9D9', fg_color: '000000' }
      }
    when "TableStyleMedium9"
      {
        border:     { color: 'FFFFFF' },
        labels:     { bg_color: '4472C4', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'B4C6E7', fg_color: '000000' },
        value_even: { bg_color: 'D9E1F2', fg_color: '000000' }
      }
    when "TableStyleMedium10"
      {
        border:     { color: 'FFFFFF' },
        labels:     { bg_color: 'ED7D31', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'F8CBAD', fg_color: '000000' },
        value_even: { bg_color: 'FCE4D6', fg_color: '000000' }
      }
    when "TableStyleMedium11"
      {
        border:     { color: 'FFFFFF' },
        labels:     { bg_color: 'A5A5A5', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'DBDBDB', fg_color: '000000' },
        value_even: { bg_color: 'EDEDED', fg_color: '000000' }
      }
    when "TableStyleMedium12"
      {
        border:     { color: 'FFFFFF' },
        labels:     { bg_color: 'FFC000', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'FFE699', fg_color: '000000' },
        value_even: { bg_color: 'FFF2CC', fg_color: '000000' }
      }
    when "TableStyleMedium13"
      {
        border:     { color: 'FFFFFF' },
        labels:     { bg_color: '5B9BD5', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'BDD7EE', fg_color: '000000' },
        value_even: { bg_color: 'DDEBF7', fg_color: '000000' }
      }
    when "TableStyleMedium14"
      {
        border:     { color: 'FFFFFF' },
        labels:     { bg_color: '70AD47', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'C6E0B4', fg_color: '000000' },
        value_even: { bg_color: 'E2EFDA', fg_color: '000000' }
      }
    else
      {
        border:     { color: 'FFFFFF' },
        labels:     { bg_color: '4472C4', fg_color: 'FFFFFF', b: true },
        value_odd:  { bg_color: 'B4C6E7', fg_color: '000000' },
        value_even: { bg_color: 'D9E1F2', fg_color: '000000' }
      }
    end
    stylize.call(style_spec)
    style_hash
  end

  def conf_total_row_styles(table_style_string = "default")
    string_parts = {
      alpha: table_style_string.match(/(?<alpha>[A-Za-z]+)/)[:alpha],
      numeric: table_style_string.match(/(?<numeric>\d+)/)[:numeric].to_i
    }
    conditions = {
      group_01: -> (string_parts) { string_parts[:alpha] == "TableStyleMedium" && (1..7).include?(string_parts[:numeric]) },
      group_02: -> (string_parts) { string_parts[:alpha] == "TableStyleMedium" && (8..14).include?(string_parts[:numeric]) }
    }
    style_hash = Hash.new do |hash, key|
      hash[key] = if conditions[:group_01].call(string_parts)
        {
          border: [
            { style: :double, edges: [:top] },
            { 
              style: :thin,
              edges: if key == :left
                [:left, :bottom]
              elsif key == :middle
                [:bottom]
              else
                [:bottom, :right]
              end
            }
          ]
        }
      elsif conditions[:group_02].call(string_parts)
        {
          border: [
            { style: :thick, edges: [:top] },
            {
              style: :thin,
              edges: if key == :left
                [:right]
              elsif key == :middle
                [:left, :right]
              else
                [:left]
              end
            }
          ]
        }
      else
        {
          border: [
            { style: :thick, edges: [:top] },
            {
              style: :thin,
              edges: if key == :left
                [:right]
              elsif key == :middle
                [:left, :right]
              else
                [:left]
              end
            }
          ]
        }
      end
    end
    [:left, :middle, :right].each { |cell_orientation| style_hash[cell_orientation] }
    style_spec = {
      "TableStyleMedium1"  => { color: '000000' },
      "TableStyleMedium2"  => { color: '8EA9DB' },
      "TableStyleMedium3"  => { color: 'F4B084' },
      "TableStyleMedium4"  => { color: 'C9C9C9' },
      "TableStyleMedium5"  => { color: 'FFD966' },
      "TableStyleMedium6"  => { color: '9BC2E6' },
      "TableStyleMedium7"  => { color: 'A9D08E' },
      "TableStyleMedium8"  => {
        cell_fill: { bg_color: '000000', fg_color: 'FFFFFF', b: true },
        border:    { color: 'FFFFFF' }
      },
      "TableStyleMedium9"  => {
        cell_fill: { bg_color: '4472C4', fg_color: 'FFFFFF', b: true },
        border:    { color: 'FFFFFF' }
      },
      "TableStyleMedium10" => {
        cell_fill: { bg_color: 'ED7D31', fg_color: 'FFFFFF', b: true },
        border:    { color: 'FFFFFF' }
      },
      "TableStyleMedium11" => {
        cell_fill: { bg_color: 'A5A5A5', fg_color: 'FFFFFF', b: true },
        border:    { color: 'FFFFFF' }
      },
      "TableStyleMedium12" => {
        cell_fill: { bg_color: 'FFC000', fg_color: 'FFFFFF', b: true },
        border:    { color: 'FFFFFF' }
      },
      "TableStyleMedium13" => {
        cell_fill: { bg_color: '5B9BD5', fg_color: 'FFFFFF', b: true },
        border:    { color: 'FFFFFF' }
      },
      "TableStyleMedium14" => {
        cell_fill: { bg_color: '70AD47', fg_color: 'FFFFFF', b: true },
        border:    { color: 'FFFFFF' }
      },
      "default"            => {
        cell_fill: { bg_color: '4472C4', fg_color: 'FFFFFF', b: true },
        border:    { color: 'FFFFFF' }
      }
    }
    stylize = {
      group_01: -> (cell_orientation, style_spec) do
        style_hash[cell_orientation][:border].each_with_index { |border, idx| style_hash[cell_orientation][:border][idx] = border.merge(style_spec) }
      end,
      group_02: -> (cell_orientation, style_spec) do
        style_hash[cell_orientation][:border].each_with_index { |border, idx| style_hash[cell_orientation][:border][idx] = border.merge(style_spec[:border]) }
        style_hash.each { |cell_orientation, style| style_hash[cell_orientation] = style_hash[cell_orientation].merge(style_spec[:cell_fill]) }
      end
    }
    if conditions[:group_01].call(string_parts)
      style_hash.each do |cell_orientation, styles|
        stylize[:group_01].call(cell_orientation, style_spec[table_style_string])
      end
    elsif conditions[:group_02].call(string_parts)
      style_hash.each do |cell_orientation, styles|
        stylize[:group_02].call(cell_orientation, style_spec[table_style_string])
      end
    else
      style_hash.each do |cell_orientation, styles|
        stylize[:group_02].call(cell_orientation, style_spec["default"])
      end
    end
    style_hash
  end

  def package
    @report
  end

  def workbook
    @workbook
  end

  def worksheet(index)
    @workbook.worksheets[index]
  end

  def serialize
    unless @report_conf["output"]["dir"].class == Array
      @report.serialize("#{@report_conf["output"]["dir"]}/#{@report_conf["output"]["name"]}.xlsx")
    else
      @report_conf["output"]["dir"].each do |dir|
        @report.serialize("#{dir}/#{@report_conf["output"]["name"]}.xlsx")
      end
    end
  end
end
