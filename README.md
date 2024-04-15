# Markdown to PowerPoint Converter

This tool is designed to convert Markdown files into PowerPoint presentations. It is especially useful for users who find thinking in Markdown more natural than traditional slide-making tools. This converter allows you to quickly transition from ideas to a structured presentation format, facilitating rapid development of slide decks directly from text-based content.

## Features

- **Single File Conversion:** Convert a specific Markdown file to a PowerPoint presentation.
- **Batch Processing:** Recursively convert all Markdown files in a directory to PowerPoint presentations.
- **Customizable Input and Output:** Specify custom paths for input and output files.
- **Automatic Directory Handling:** Automatically creates necessary output directories.

## Installation

Ensure that Python and pip are installed on your system, then install the required Python package:

```
bash
pip install python-pptx
```

## Usage
#### Basic Usage
To convert a single Markdown file to a PowerPoint presentation using the default paths (input/slides.md to output/slides.pptx):

```python app.py ```

This assumes that the input file is located in the input directory and has the .md extension. The output file will be saved in the output directory with the .pptx extension.


#### Custom File Conversion
To convert a specific Markdown file to a PowerPoint file with custom paths:
```python app.py -f path/to/your/markdown.md -o path/to/your/output.pptx```


#### Recursive Batch Conversion
To recursively process all Markdown files in the input directory, outputting PowerPoint files to the output directory:
```python script_name.py -i path/to/input/directory -o path/to/output/directory```

#### Benefits
Using this tool enhances productivity by enabling quick conversions of structured Markdown into visually appealing presentations. It is ideal for people like me, who prefer to outline and draft content in a code editor but sometimes need to deliver final products in a graphical format suitable for professional or educational presentations. I think better in Markdown than in PowerPoint.

