require "powerpoint/version"
require 'powerpoint/util'
require 'powerpoint/slide/intro'
require 'powerpoint/slide/textual'
require 'powerpoint/slide/pictorial'
require 'powerpoint/slide/text_picture_split'
require 'powerpoint/slide/picture_description'
require 'powerpoint/slide/picture_without_title'
require 'powerpoint/slide/fluvipSlide/main'
require 'powerpoint/slide/fluvipSlide/basic'
require 'powerpoint/slide/fluvipSlide/just_background'
require 'powerpoint/compression'
require 'powerpoint/presentation'

module Powerpoint
  ROOT_PATH = File.expand_path("../..", __FILE__)
  TEMPLATE_PATH = "#{ROOT_PATH}/template/"
  VIEW_PATH = "#{ROOT_PATH}/lib/powerpoint/views/"
end
