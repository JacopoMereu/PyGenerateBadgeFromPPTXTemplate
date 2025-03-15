# Badge Generator

This script generates conference badges using data from a CSV file and a PowerPoint template. It creates a new PowerPoint file with personalized badges for each attendee.

## Features
- Reads attendee data from a CSV file or uses hardcoded sample data.
- Uses a PowerPoint template to generate badge slides.
- Dynamically fills placeholders with attendee information.
- Ensures the output file exists before saving.

## Requirements
- Python 3.x
- `pandas`
- `python-pptx`

## Installation
Install dependencies using pip:
```sh
pip install pandas python-pptx
```

## Usage

1. Duplicate one of the sample folders (e.g., `EICS24`). Each folder contains three files:
   - `badge.pptx`: The source file with shapes and text used to create the badge. These elements are converted into an image through a screenshot (using the Stamp key for simplicity).
   - `badge-layout.pptx`: The layout file. Open `View > Slide Master` and replace the background of all images with the screenshot taken in the previous step.
   - `OUTPUT.pptx`: A sample output file (this file is not important).
2. Adapt `badge-layout.pptx` as needed.
3. Configure the input parameters in the script.
4. Run the `main.py` script. It will overwrite the `OUTPUT.pptx` file with the new badges.

## License
This project is licensed under the MIT License.