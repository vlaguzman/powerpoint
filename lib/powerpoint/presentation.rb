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
      @slides << Powerpoint::Slide::FluvipSlide::Main.new(presentation: self, title: title, image_path: image_path)
    end

    def add_fluvip_just_background_slide(title, background_path)
      @slides << Powerpoint::Slide::FluvipSlide::JustBackground.new(presentation: self, title: title, background_path: background_path)
    end

    def add_fluvip_contries_slide(title)
      @slides << Powerpoint::Slide::FluvipSlide::Basic.new(presentation: self, title: title, rels_path: 'fluvip/contries_rels.xml.erb', slide_path: 'fluvip/contries_slide.xml.erb')
    end

    def add_fluvip_description_slide(title)
      @slides << Powerpoint::Slide::FluvipSlide::Basic.new(presentation: self, title: title, rels_path: 'fluvip/description_rels.xml.erb', slide_path: 'fluvip/description_slide.xml.erb')
    end

    def add_fluvip_objective_slide(title)
      @slides << Powerpoint::Slide::FluvipSlide::Basic.new(presentation: self, title: title, rels_path: 'fluvip/objective_rels.xml.erb', slide_path: 'fluvip/objective_slide.xml.erb')
    end

    def add_fluvip_info_campaign_slide(title)
      @slides << Powerpoint::Slide::FluvipSlide::Basic.new(presentation: self, title: title, rels_path: 'fluvip/info_campaign_rels.xml.erb', slide_path: 'fluvip/info_campaign_slide.xml.erb')
    end

    def add_fluvip_info_influencer_slide(title)
      @slides << Powerpoint::Slide::FluvipSlide::Basic.new(presentation: self, title: title, rels_path: 'fluvip/info_influencer_rels.xml.erb', slide_path: 'fluvip/info_influencer_slide.xml.erb')
    end

    def add_fluvip_campaign_resume_slide(title)
      @slides << Powerpoint::Slide::FluvipSlide::Basic.new(presentation: self, title: title, rels_path: 'fluvip/campaign_resume_rels.xml.erb', slide_path: 'fluvip/campaign_resume_slide.xml.erb')
    end

    def add_fluvip_signature_slide(title)
      @slides << Powerpoint::Slide::FluvipSlide::Basic.new(presentation: self, title: title, rels_path: 'fluvip/signature_rels.xml.erb', slide_path: 'fluvip/signature_slide.xml.erb')
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