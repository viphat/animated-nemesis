class Code
  attr_accessor :code, :label, :index, :bold

  def initialize(code,label,index=0)
    @code, @label, @index = code,label,index
    @bold = false
  end

end