require 'fileutils'
require 'erb'

module Powerpoint
  module Slide
    module FluvipSlide
      class Main
        include Powerpoint::Util
        
        attr_reader :title

        def initialize(options={})
          require_arguments [:title], options
          options.each {|k, v| instance_variable_set("@#{k}", v)}
        end

        def save(extract_path, index)
          save_rel_xml(extract_path, index)
          save_slide_xml(extract_path, index)
        end

        def save_rel_xml(extract_path, index)
          render_view('fluvip/main_rels.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
        end
        private :save_rel_xml

        def save_slide_xml(extract_path, index)
          render_view('fluvip/main_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
        end
        private :save_slide_xml
      end
    end
  end
end