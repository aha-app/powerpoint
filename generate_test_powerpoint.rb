require 'bundler/inline'

gemfile do
  source 'https://rubygems.org'
  gem 'powerpoint', path: '.' # Assuming we're in the gem's root directory
end

require 'tempfile'

def pixle_to_pt(px)
  px * 12700
end

def create_test_powerpoint
  # Create a new presentation
  deck = Powerpoint::Presentation.new

  # Add title slide
  deck.add_intro("Test Presentation", "Created with Ruby")

  # Add text-only slide
  deck.add_textual_slide(
    "Text Slide",
    ["First bullet point", "Second bullet point", "Third bullet point"]
  )

  # Test images with their dimensions
  image_files = [
    {
      path: "samples/images/landscape.jpeg",  # Landscape 16:9
      name: "Landscape Image",
      width: 1920,
      height: 1080
    },
    {
      path: "samples/images/3_4.jpeg",  # Portrait 3:4
      name: "Portrait Image",
      width: 900,
      height: 1200
    },
    {
      path: "samples/images/square.jpeg",  # Square 1:1
      name: "Square Image",
      width: 1000,
      height: 1000
    }
  ]

  # Add image slides
  slide_width = pixle_to_pt(720)
  default_width = pixle_to_pt(550)
  default_height = pixle_to_pt(400)

  image_files.each do |image|
    next unless File.exist?(image[:path])

    image_width = pixle_to_pt(image[:width])
    image_height = pixle_to_pt(image[:height])

    # Calculate dimensions
    if image_height > image_width && image_height > default_height
      new_height = default_height
      ratio = new_height / image_height.to_f
      new_width = (image_width.to_f * ratio).round
    else
      new_width = default_width < image_width ? default_width : image_width
      ratio = new_width / image_width.to_f
      new_height = (image_height.to_f * ratio).round
    end

    coords = {
      x: (slide_width / 2) - (new_width / 2),
      y: pixle_to_pt(120),
      cx: new_width,
      cy: new_height
    }

    # Add different types of slides with images
    deck.add_pictorial_slide(
      "#{image[:name]} - Full",
      image[:path],
      coords
    )

    deck.add_text_picture_slide(
      "#{image[:name]} - With Text",
      image[:path],
      ["Left side text", "With bullet points", "Next to image"]
    )

    deck.add_picture_description_slide(
      "#{image[:name]} - With Description",
      image[:path],
      ["Description below the image", "With multiple points", "For detail"]
    )
  end

  # Save the presentation
  output_dir = "output"
  output_file = "test_presentation"
  
  begin
    # Ensure output directory exists
    FileUtils.mkdir_p(output_dir)
    
    # Save with base name (PowerPoint gem will append timestamp)
    saved_path = deck.save("#{output_dir}/#{output_file}")
    
    # Rename the file to remove timestamp
    final_path = "#{output_dir}/#{output_file}.pptx"
    FileUtils.mv(saved_path, final_path)
    
    puts "Presentation saved to: #{final_path}"
  rescue StandardError => e
    puts "Error saving presentation: #{e.message}"
    puts "Stack trace:"
    puts e.backtrace
    exit 1
  end
end

# Run the test
create_test_powerpoint

