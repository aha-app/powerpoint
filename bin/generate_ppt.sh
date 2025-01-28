#!/bin/bash

# Exit on error
set -e

ruby ./generate_test_powerpoint.rb

# Get the most recent pptx file in the output directory
latest_ppt=$(ls -t output/*.pptx | head -n1)

if [ -z "$latest_ppt" ]; then
    echo "No PPTX files found in output directory"
    exit 1
fi

echo "Processing: $latest_ppt"

# Remove existing unzipped directory if it exists
rm -rf output/unzipped

# Create unzipped directory
mkdir -p output/unzipped

# Unzip the PPTX file
unzip "$latest_ppt" -d output/unzipped

echo "Successfully unzipped PPTX to output/unzipped/"
