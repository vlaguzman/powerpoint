require 'fileutils'
require 'erb'

module Powerpoint
  module Slide
    module FluvipSlide
      class InfoSocialNetworkAccount
        include Powerpoint::Util
        
        attr_reader :title, :image_path, :profile_image_name, :provider_path, :provider_name, :coords, :influencer_data, :gender_data, :ages_data, 
                    :interests_data, :countries_data, :cities_data, :flag_file_name, :flag_file_name2, :flag_file_name3, :flag_file_name4, 
                    :flag_file_name5, :top_followers_data, :sn_image_name

        def initialize(options={})
          require_arguments [:title, :image_path, :provider_path, :influencer_data, :gender_data, :ages_data, :interests_data, :countries_data], options
          options.each {|k, v| instance_variable_set("@#{k}", v)}

          @provider_name = File.basename(@provider_path)
          @sn_image_name = File.basename(@sn_image_path)
          @flag_file_name = File.basename(@countries_data[0][:image_path])
          @flag_file_name2 = File.basename(@countries_data[1][:image_path])
          @flag_file_name3 = File.basename(@countries_data[2][:image_path])
          @flag_file_name4 = File.basename(@countries_data[3][:image_path])
          @profile_image_name = File.basename(@image_path)
          assign_top_follower_file_names if @top_followers_data
        end

        def save(extract_path, index)
          begin
            copy_media(extract_path, @image_path)
            copy_media(extract_path, @provider_path)
            copy_media(extract_path, @sn_image_path)
            copy_media(extract_path, @countries_data[0][:image_path])
            copy_media(extract_path, @countries_data[1][:image_path])
            copy_media(extract_path, @countries_data[2][:image_path])
            copy_media(extract_path, @countries_data[3][:image_path])
            copy_top_followers_media_data(extract_path) if @top_followers_data
          rescue StandardError => error
            Rails.logger.error error.message
          end
          save_rel_xml(extract_path, index)
          save_slide_xml(extract_path, index)
        end

        def assign_top_follower_file_names
          @top_follower_file_name = File.basename(@top_followers_data[0][:picture_url])
          @top_follower_file_name2 = File.basename(@top_followers_data[1][:picture_url])
          @top_follower_file_name3 = File.basename(@top_followers_data[2][:picture_url])
          @top_follower_file_name4 = File.basename(@top_followers_data[3][:picture_url])
          @top_follower_file_name5 = File.basename(@top_followers_data[4][:picture_url])
        end

        def copy_top_followers_media_data(extract_path)
          copy_media(extract_path, @top_followers_data[0][:picture_url])
          copy_media(extract_path, @top_followers_data[1][:picture_url])
          copy_media(extract_path, @top_followers_data[2][:picture_url])
          copy_media(extract_path, @top_followers_data[3][:picture_url])
          copy_media(extract_path, @top_followers_data[4][:picture_url])
        end
        
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
