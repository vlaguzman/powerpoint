require 'zip/filesystem'
require 'fileutils'
require 'tmpdir'

module Powerpoint
  class Presentation
    include Powerpoint::Util

    attr_reader :slides

    def initialize
      @slides = []
    end

    def self.new_with_providers(providers)
      @providers = providers
      self.new
    end

    def add_intro(title, subtitile = nil)
      existing_intro_slide = @slides.select {|s| s.class == Powerpoint::Slide::Intro}[0]
      slide = Powerpoint::Slide::Intro.new(presentation: self, title: title, subtitile: subtitile)
      if existing_intro_slide
        @slides[@slides.index(existing_intro_slide)] = slide 
      else
        @slides.insert 0, slide
      end
    end

    def add_textual_slide(title, content = [])
      @slides << Powerpoint::Slide::Textual.new(presentation: self, title: title, content: content)
    end

    def add_pictorial_slide(title, image_path, coords = {})
      @slides << Powerpoint::Slide::Pictorial.new(presentation: self, title: title, image_path: image_path, coords: coords)
    end

    def add_text_picture_slide(title, image_path, content = [])
      @slides << Powerpoint::Slide::TextPicSplit.new(presentation: self, title: title, image_path: image_path, content: content)
    end

    def add_picture_description_slide(title, image_path, content = [])
      @slides << Powerpoint::Slide::DescriptionPic.new(presentation: self, title: title, image_path: image_path, content: content)
    end

    def add_picture_without_title_slide(title, image_path)
      @slides << Powerpoint::Slide::PictureWithoutTitle.new(presentation: self, title: title, image_path: image_path)
    end

    def add_background_picture_slide(title, background_path, image_path)
      @slides << Powerpoint::Slide::BackgroundPicture.new(presentation: self, title: title, background_path: background_path, image_path: image_path)
    end

    def add_fluvip_main_slide(title, image_path)
      @slides << Powerpoint::Slide::FluvipSlide::Main.new(presentation: self, title: title, image_path: image_path, rels_path: 'fluvip/main_rels.xml.erb', slide_path: 'fluvip/main_slide.xml.erb')
    end

    def add_fluvip_main_social_network_slide(title, image_path)
      @slides << Powerpoint::Slide::FluvipSlide::Main.new(presentation: self, title: title, image_path: image_path, rels_path: 'fluvip/main_social_network_rels.xml.erb', slide_path: 'fluvip/main_social_network_slide.xml.erb')
    end

    def add_fluvip_just_background_slide(title, background_path)
      @slides << Powerpoint::Slide::FluvipSlide::JustBackground.new(presentation: self, title: title, background_path: background_path)
    end

    def add_fluvip_text_slide(title, text, rels_path, slide_path)
      @slides << Powerpoint::Slide::FluvipSlide::Text.new(presentation: self, title: title, text: text, rels_path: rels_path, slide_path: slide_path)
    end

    def add_fluvip_contries_slide(title, text)
      add_fluvip_text_slide(title, text, 'fluvip/contries_rels.xml.erb', 'fluvip/contries_slide.xml.erb')
    end

    def add_fluvip_description_slide(title, text)
      add_fluvip_text_slide(title, text, 'fluvip/description_rels.xml.erb', 'fluvip/description_slide.xml.erb')
    end

    def add_fluvip_signature_slide(title, text)
      add_fluvip_text_slide(title, text, 'fluvip/signature_rels.xml.erb', 'fluvip/signature_slide.xml.erb')
    end

    def add_fluvip_objective_slide(title, objective)
      @slides << Powerpoint::Slide::FluvipSlide::Objective.new(presentation: self, title: title, objective: objective, rels_path: 'fluvip/objective_rels.xml.erb', slide_path: 'fluvip/objective_slide.xml.erb')
    end

    def add_fluvip_info_campaign_slide(title, objective, creation_date, image_path, influencers_data, posts_data, users_data, impressions_data)
      @slides << Powerpoint::Slide::FluvipSlide::InfoCampaign.new(presentation: self, title: title, objective: objective, creation_date: creation_date, image_path: image_path, influencers_data: influencers_data, posts_data: posts_data, users_data: users_data, impressions_data: impressions_data, providers: @providers, rels_path: 'fluvip/info_campaign_rels.xml.erb', slide_path: 'fluvip/info_campaign_slide.xml.erb')
    end

    def add_fluvip_complete_info_social_network_account_slide(title, image_path, provider_path, influencer_data, gender_data, ages_data, interests_data, countries_data, cities_data, top_followers_data, sn_image_path)
      @slides << Powerpoint::Slide::FluvipSlide::InfoSocialNetworkAccount.new(presentation: self, title: title, image_path: image_path, provider_path: provider_path, influencer_data: influencer_data, gender_data: gender_data, ages_data: ages_data, interests_data: interests_data, countries_data: countries_data, cities_data: cities_data, top_followers_data: top_followers_data, sn_image_path: sn_image_path, rels_path: 'fluvip/info_influencer_rels.xml.erb', slide_path: 'fluvip/info_influencer_slide.xml.erb')
    end

    def add_fluvip_info_social_network_account_slide(title, image_path, provider_path, influencer_data, gender_data, ages_data, interests_data, countries_data, sn_image_path)
      @slides << Powerpoint::Slide::FluvipSlide::InfoSocialNetworkAccount.new(presentation: self, title: title, image_path: image_path, provider_path: provider_path, influencer_data: influencer_data, gender_data: gender_data, ages_data: ages_data, interests_data: interests_data, countries_data: countries_data, sn_image_path: sn_image_path, rels_path: 'fluvip/info_social_network_rels.xml.erb', slide_path: 'fluvip/info_social_network_slide.xml.erb')
    end

    def add_fluvip_totals_campaing_slide(title, influencers_data, publications_data, reach_data, impressions_data, price_data)
      @slides << Powerpoint::Slide::FluvipSlide::TotalsCampaign.new(presentation: self, title: title, influencers_data: influencers_data, publications_data: publications_data, reach_data: reach_data, impressions_data: impressions_data, price_data: price_data, rels_path: 'fluvip/totals_campaign_rels.xml.erb', slide_path: 'fluvip/totals_campaign_slide.xml.erb')
    end

    def add_fluvip_campaign_resume_slide(title)
      @slides << Powerpoint::Slide::FluvipSlide::Basic.new(presentation: self, title: title, rels_path: 'fluvip/campaign_resume_rels.xml.erb', slide_path: 'fluvip/campaign_resume_slide.xml.erb')
    end

    def add_fluvip_flags_slide(title)
      @slides << Powerpoint::Slide::FluvipSlide::Basic.new(presentation: self, title: title, rels_path: 'fluvip/flags_rels.xml.erb', slide_path: 'fluvip/flags_slide.xml.erb')
    end

    def save(path)
      Dir.mktmpdir do |dir|
        extract_path = "#{dir}/extract_#{Time.now.strftime("%Y-%m-%d-%H%M%S")}"

        # Copy template to temp path
        FileUtils.copy_entry(TEMPLATE_PATH, extract_path)

        # Remove keep files
        Dir.glob("#{extract_path}/**/.keep").each do |keep_file|
          FileUtils.rm_rf(keep_file)
        end

        # Render/save generic stuff
        render_view('content_type.xml.erb', "#{extract_path}/[Content_Types].xml")
        render_view('presentation.xml.rel.erb', "#{extract_path}/ppt/_rels/presentation.xml.rels")
        render_view('presentation.xml.erb', "#{extract_path}/ppt/presentation.xml")
        render_view('app.xml.erb', "#{extract_path}/docProps/app.xml")

        # Save slides
        slides.each_with_index do |slide, index|
          slide.save(extract_path, index + 1)
        end

        # Create .pptx file
        File.delete(path) if File.exist?(path)
        Powerpoint.compress_pptx(extract_path, path)
      end

      path
    end

    def file_types
      slides.map {|slide| slide.file_type if slide.respond_to? :file_type }.compact.uniq
    end
  end
end