module Powerpoint
  module Slide
    class Base
      include Powerpoint::Util

      def layout_name
        'content' # Default layout for most slides
      end

      protected

      def relationship_xml_path(extract_path, index)
        "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels"
      end

      def slide_xml_path(extract_path, index)
        "#{extract_path}/ppt/slides/slide#{index}.xml"
      end
    end
  end
end
