class Code
  attr_accessor :code, :label, :index, :bold, :sub_label

  def initialize(code,label,index=0,sub_label=nil)
    @code, @label, @index,@sub_label = code,label, index, sub_label
    @bold = false
  end

end