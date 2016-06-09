require 'fileutils'
require 'erb'

module Powerpoint
  module Slide
    module FluvipSlide
      class TotalsCampaign < Basic
        include Powerpoint::Util
        
        attr_reader :title, :influencers_data, :publications_data, :reach_data, :impressions_data, :price_data

        def initialize(options={})
          require_arguments [:title, :influencers_data, :publications_data, :reach_data, :impressions_data, :price_data], options
          options.each {|k, v| instance_variable_set("@#{k}", v)}
        end
        
      end
    end
  end
end