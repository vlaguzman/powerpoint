require 'fileutils'
require 'erb'

module Powerpoint
  module Slide
    module FluvipSlide
      class JustBackground
        include Powerpoint::Util
        
        attr_reader :title, :background_name, :background_path

        def initialize(options={})
          require_arguments [:title, :background_path], options
          options.each {|k, v| instance_variable_set("@#{k}", v)}
          @background_name = File.basename(@background_path)
        end

        def save(extract_path, index)
          copy_media(extract_path, @background_path)
          save_rel_xml(extract_path, index)
          save_slide_xml(extract_path, index)
        end

        def save_rel_xml(extract_path, index)
          render_view('fluvip/just_background_rels.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
        end
        private :save_rel_xml

        def save_slide_xml(extract_path, index)
          render_view('fluvip/just_background_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
        end
        private :save_slide_xml
      end
    end
  end
end