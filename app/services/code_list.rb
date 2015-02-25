class CodeList
  attr_accessor :question, :filter, :qbegin, :qend
  def initialize(q='',filter='',qbegin=nil,qend=nil)
    @question, @filter, @qbegin, @qend = q,filter,qbegin,qend
  end
end