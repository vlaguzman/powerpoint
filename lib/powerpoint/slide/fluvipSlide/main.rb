require 'fileutils'
require 'erb'

module Powerpoint
  module Slide
    module FluvipSlide
      class Main
        include Powerpoint::Util
        
        attr_reader :title, :image_path, :image_name, :coords

        def initialize(options={})
          require_arguments [:title, :image_path], options
          options.each {|k, v| instance_variable_set("@#{k}", v)}
          @image_name = File.basename(@image_path)
          @coords = default_coords
        end

        def save(extract_path, index)
          copy_media(extract_path, @image_path)
          save_rel_xml(extract_path, index)
          save_slide_xml(extract_path, index)
        end

        def default_coords
          dimensions = FastImage.size(image_path)
          default_width = default_height = pixle_to_pt(150)
          image_width, image_height = dimensions.map {|d| pixle_to_pt(d)}
          new_width = default_width < image_width ? default_width : image_width
          ratio = new_width / image_width.to_f
          new_height = (image_height.to_f * ratio).round

          {cx: new_width, cy: new_height}
        end
        private :default_coords


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