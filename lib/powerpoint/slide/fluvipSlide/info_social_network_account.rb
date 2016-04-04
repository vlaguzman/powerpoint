require 'fileutils'
require 'erb'

module Powerpoint
  module Slide
    module FluvipSlide
      class InfoSocialNetworkAccount
        include Powerpoint::Util
        
        attr_reader :title, :image_path, :image_name, :provider_path, :provider_name, :coords, :influencer_data, :gender_data, :ages_data, 
                    :interests_data, :countries_data, :flag_file_name, :flag_file_name2, :flag_file_name3, :flag_file_name4, 
                    :flag_file_name5

        def initialize(options={})
          require_arguments [:title, :image_path, :provider_path, :influencer_data, :gender_data, :ages_data, :interests_data, :countries_data], options
          options.each {|k, v| instance_variable_set("@#{k}", v)}
          @image_name = File.basename(@image_path)
          @provider_name = File.basename(@provider_path)
          @flag_file_name = File.basename(countries_data[0][:image_path])
          @flag_file_name2 = File.basename(countries_data[1][:image_path])
          @flag_file_name3 = File.basename(countries_data[2][:image_path])
          @flag_file_name4 = File.basename(countries_data[3][:image_path])
          @flag_file_name5 = File.basename(countries_data[4][:image_path])
          @coords = default_coords
        end

        def save(extract_path, index)
          copy_media(extract_path, @image_path)
          copy_media(extract_path, @provider_path)
          copy_media(extract_path, @countries_data[0][:image_path])
          copy_media(extract_path, @countries_data[1][:image_path])
          copy_media(extract_path, @countries_data[2][:image_path])
          copy_media(extract_path, @countries_data[3][:image_path])
          copy_media(extract_path, @countries_data[4][:image_path])                             
          save_rel_xml(extract_path, index)
          save_slide_xml(extract_path, index)
        end

        def default_coords
          default_width = default_height = pixle_to_pt(100)

          dimensions = FastImage.size(image_path)
          image_width, image_height = dimensions.map {|d| pixle_to_pt(d)}
          new_width = default_width < image_width ? default_width : image_width
          ratio = new_width / image_width.to_f
          new_height = (image_height.to_f * ratio).round

          {cx: new_width, cy: new_height}
        end
        private :default_coords

        def save_rel_xml(extract_path, index)
          render_view(@rels_path, "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
        end
        private :save_rel_xml

        def save_slide_xml(extract_path, index)
          render_view(@slide_path, "#{extract_path}/ppt/slides/slide#{index}.xml")
        end
        private :save_slide_xml

      end
    end
  end
end