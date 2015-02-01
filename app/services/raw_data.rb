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
                :sheet_name

  def initialize
    @wtd_resp = @resp = @question = @base = @table_name = @filters = @means = @medians = @mode = @std_deviation = @totals_count = @totals_percent = @header_label = @header = @sheet_name = nil
    @val = []
  end
end