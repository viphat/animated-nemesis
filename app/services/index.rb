# https://netguru.co/blog/service-objects-in-rails-will-help
# http://www.tutorialspoint.com/ruby/ruby_object_oriented.htm

class Index
  attr_accessor :index,
                :question,
                :sheet_name,
                :link,
                :filter,
                :error

  def initialize(i=nil,q=nil,s=nil,l=nil,f=nil,e=nil)
    @index, @question, @sheet_name, @link, @filter, @error = i,q,s,l,f,e
  end

end