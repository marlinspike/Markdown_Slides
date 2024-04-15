import argparse
import os
from glob import glob
from pptx import Presentation

def parse_markdown(md_file):
    """
    Parses markdown file to slides and bullet points.
    """
    slides = []
    current_slide = None
    with open(md_file, 'r') as file:
        md_text = file.read()

    for line in md_text.strip().split('\n'):
        if line.startswith('Slide'):
            if current_slide:
                slides.append(current_slide)
            title = line.split(': ', 1)[1] if ': ' in line else line
            current_slide = {'title': title, 'content': []}
        elif line.startswith('- '):
            current_slide['content'].append((line[2:], 'main'))
        elif line.startswith('  - '):
            current_slide['content'].append((line[4:], 'sub'))

    if current_slide:
        slides.append(current_slide)
    return slides

def create_presentation(slides, pptx_file):
    """
    Creates a PowerPoint presentation from parsed slides.
    """
    prs = Presentation()
    bullet_slide_layout = prs.slide_layouts[1]
    
    for slide_info in slides:
        slide = prs.slides.add_slide(bullet_slide_layout)
        title, content = slide_info['title'], slide_info['content']
        
        title_placeholder = slide.shapes.title
        title_placeholder.text = title
        
        content_placeholder = slide.placeholders[1]
        for text, level in content:
            p = content_placeholder.text_frame.add_paragraph()
            p.text = text
            p.level = 0 if level == 'main' else 1

    prs.save(pptx_file)
    print(f"Presentation saved as {pptx_file}")

def process_files(input_path, output_path, recursive):
    """
    Processes files from input path and saves them to output path.
    If recursive is True, processes all .md files in the directory recursively.
    """
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    if recursive:
        pattern = '**/*.md' if recursive else '*.md'
        md_files = glob(os.path.join(input_path, pattern), recursive=recursive)
    else:
        md_files = [input_path]

    for md_file in md_files:
        pptx_file = os.path.join(output_path, os.path.basename(md_file).replace('.md', '.pptx'))
        slides = parse_markdown(md_file)
        create_presentation(slides, pptx_file)

def main():
    parser = argparse.ArgumentParser(description="Convert Markdown files to PowerPoint presentations")
    parser.add_argument('-f', '--file', type=str, default=os.path.join('input', 'slides.md'),
                        help='Input markdown file name, default is "input/slides.md"')
    parser.add_argument('-o', '--output', type=str, default=os.path.join('output', 'slides.pptx'),
                        help='Output PowerPoint file name, default is "output/slides.pptx"')
    parser.add_argument('-r', '--recursive', action='store_true',
                        help='Recursively process all markdown files in the input directory')
    args = parser.parse_args()

    if args.recursive:
        input_path = os.path.dirname(args.file) if os.path.dirname(args.file) else 'input'
        output_path = 'output'
        process_files(input_path, output_path, args.recursive)
    else:
        if not os.path.exists(os.path.dirname(args.output)):
            os.makedirs(os.path.dirname(args.output))
        slides = parse_markdown(args.file)
        create_presentation(slides, args.output)

if __name__ == '__main__':
    main()
