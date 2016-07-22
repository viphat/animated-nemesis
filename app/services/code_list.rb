class CodeList
  attr_accessor :question, :filters, :qbegin, :qend, :code
  def initialize(q='',filters='',qbegin=nil,qend=nil)
    @question, @filters, @qbegin, @qend = q, filters, qbegin, qend
    @code = []
  end
end