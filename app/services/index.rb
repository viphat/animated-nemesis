# https://netguru.co/blog/service-objects-in-rails-will-help
# http://www.tutorialspoint.com/ruby/ruby_object_oriented.htm

class Index
  attr_accessor :index,
                :question,
                :sheet_name,
                :link,
                :filter,
                :error,
                :means

  def initialize(i=nil,q=nil,s=nil,l=nil,f=nil,e=nil,m=nil)
    @index, @question, @sheet_name, @link, @filter, @error, @means = i,q,s,l,f,e,m
  end

  def build!(hash)
    @index, @question, @sheet_name, @link, @filter, @error, @means = hash["index"], hash["question"], hash["sheet_name"], hash["link"], hash["filter"], hash["error"], hash["means"]
    self
  end

end