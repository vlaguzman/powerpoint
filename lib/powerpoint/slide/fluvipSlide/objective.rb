require 'fileutils'
require 'erb'

module Powerpoint
  module Slide
    module FluvipSlide
      class Objective < Basic
        include Powerpoint::Util
        
        attr_reader :objective

        def initialize(options={})
          require_arguments [:title, :objective], options
          options.each {|k, v| instance_variable_set("@#{k}", v)}
        end

      end
    end
  end
end