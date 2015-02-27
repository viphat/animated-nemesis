class RawData
  attr_accessor :wtd_resp,
                :resp,
                :question,
                :base,
                :table_name,
                :filters,
                :means,
                :medians,
                :mode,
                :std_deviation,
                :totals_count,
                :totals_percent,
                :header_label,
                :header,
                :val,
                :sheet_name,
                :codelist

  def initialize
    @wtd_resp = @resp = @question = @base = @table_name = @filters = @means = @medians = @mode = @std_deviation = @totals_count = @totals_percent = @header_label = @header = @sheet_name = @codelist = nil
    @val = []
  end

  def build!(hash)
    @wtd_resp, @resp = hash['wtd_resp'],hash['resp']
    @question = hash['question']
    @base = hash['base']
    @table_name = hash['table_name']
    @filters = hash['filters']
    @means, @mode, @medians = hash['means'], hash['mode'], hash['medians']
    @std_deviation = hash['std_deviation']
    @totals_count, @totals_percent = hash['totals_count'], hash['totals_percent']
    @header_label, @header = hash['header_label'], hash['header']
    @sheet_name = hash['sheet_name']
    @val = hash['val']
    @codelist = nil
  end

end