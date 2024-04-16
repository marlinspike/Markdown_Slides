# Markdown to PowerPoint Converter

This tool is designed to convert Markdown files into PowerPoint presentations. It is especially useful for people who find thinking in Markdown more natural than traditional slide-making tools. This converter allows you to quickly transition from ideas to a structured presentation format, facilitating rapid development of slide decks directly from text-based content.

## Features

- **Single File Conversion:** Convert a specific Markdown file to a PowerPoint presentation.
- **Batch Processing:** Recursively convert all Markdown files in a directory to PowerPoint presentations.
- **Customizable Input and Output:** Specify custom paths for input and output files.
- **Automatic Directory Handling:** Automatically creates necessary output directories.
- **Visual Indentation:** Indentation levels in the Markdown are visually represented in the PowerPoint slides to reflect the structure of the content.
- **Configurable Text Formatting:** Configure text formatting such as bold for specific levels via environment settings.
- **Markdown Syntax Support:** Supports a wide range of Markdown syntax, including:
  - Headings (#, ##, ###, ####, #####)
  - Italics and Bold (*, _)
  - Underline (__)
  - Strikethrough (~~)
  - Inline Code (`)
  - Hyperlinks (text)
  - Images (![alt text](image url))

## Installation

Ensure that Python and pip are installed on your system, then install the required Python package:

```bash
pip install -r requirements.txt
```

## Configuration
Adjust the .env file to set font sizes and bold formatting for different indentation levels:

```
level-one=18
level-two=16
level-three=14
bold-level-one=true  # Set to 'false' if you do not want level-1 text to be bold
```

## Usage
#### Basic Usage
To convert a single Markdown file to a PowerPoint presentation using the default paths (input/slides.md to output/slides.pptx):

```python app.py ```

This assumes that the input file is located in the input directory and has the .md extension. The output file will be saved in the output directory with the .pptx extension.


#### Custom File Conversion
To convert a specific Markdown file to a PowerPoint file with custom paths:
```python app.py -f path/to/your/markdown.md```

The output file will be saved in the output directory with the same name as the input file (with the .pptx extension).

#### Custom Output File
To specify a custom output file name when converting a single Markdown file:

````python app.py -f path/to/your/markdown.md -o path/to/your/output.pptx````

#### Recursive Batch Conversion
To recursively process all Markdown files in the input directory, outputting PowerPoint files to the output directory:
```python script_name.py -r```

#### Give it a whirl
Test it out immediately with a demo file:
```python app.py -f input/demo.md```
Or try it recursively on the demo directory:
```python app.py -r```

### Tips for Usage
- Ensure all indentation in your Markdown files uses consistent spacing as expected by the script.
- Customize the .env file to fit your presentation style preferences regarding font sizes and text attributes like bold.
- When using Markdown syntax, make sure to follow the proper formatting to ensure correct rendering in the PowerPoint presentation. Refer to the supported Markdown syntax in the Features section.