require 'fileutils'
require 'erb'

module Powerpoint
  module Slide
    module FluvipSlide
      class Text < Basic
        include Powerpoint::Util
        
        attr_reader :text

        def initialize(options={})
          require_arguments [:title, :text], options
          options.each {|k, v| instance_variable_set("@#{k}", v)}
        end

      end
    end
  end
end