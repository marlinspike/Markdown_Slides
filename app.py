from pptx.util import Pt
from pptx import Presentation
import os
import re
import argparse
from glob import glob
from dotenv import load_dotenv

# Load environment settings for font sizes and bold setting
load_dotenv()
font_sizes = {
    'level-1': os.getenv('level-one', '18'),
    'level-2': os.getenv('level-two', '16'),
    'level-3': os.getenv('level-three', '14')
}
is_bold_level_one = os.getenv('bold-level-one', 'true').lower() == 'true'  # Default to bold for level-1

def parse_markdown(md_file):
    """
    Parses markdown file to slides and bullet points, considering indentation levels.
    Uses 2 spaces per indentation level.
    """
    slides = {}
    with open(md_file, 'r') as file:
        md_text = file.read()

    current_slide_number = None
    for line in md_text.strip().split('\n'):
        slide_match = re.match(r'^Slide (\d+): (.+)', line)
        if slide_match:
            current_slide_number = slide_match.group(1)
            if current_slide_number not in slides:
                slides[current_slide_number] = {'title': slide_match.group(2), 'content': []}
        elif current_slide_number:
            leading_spaces = len(line) - len(line.lstrip(' '))
            level = leading_spaces // 2
            text = line.strip()
            if text.startswith('- '):
                text = text[2:]
            slides[current_slide_number]['content'].append((text, level))

    return list(slides.values())

def get_font_size(level):
    """
    Fetches the font size from the environment settings based on the indentation level.
    """
    key = f'level-{level+1}'
    font_size = font_sizes.get(key, '12')
    return Pt(int(font_size))

def create_presentation(slides, pptx_file):
    prs = Presentation()
    bullet_slide_layout = prs.slide_layouts[1]  # 'Title and Content' layout

    for slide_info in slides:
        slide = prs.slides.add_slide(bullet_slide_layout)
        title_placeholder = slide.shapes.title
        title_placeholder.text = slide_info['title']

        content_placeholder = slide.placeholders[1]
        for text, level in slide_info['content']:
            p = content_placeholder.text_frame.add_paragraph()
            p.text = text
            p.level = level  # Apply visual indentation
            p.font.size = get_font_size(level)
            if level == 0 and is_bold_level_one:  # Apply bold to level-1 if configured
                p.font.bold = True

    prs.save(pptx_file)
    print(f"Presentation saved as {pptx_file}")

def main():
    parser = argparse.ArgumentParser(description="Convert Markdown files to PowerPoint presentations")
    parser.add_argument('-f', '--file', type=str, default='input/slides.md',
                        help='Input markdown file name, default is "input/slides.md"')
    parser.add_argument('-o', '--output', type=str, default='output/slides.pptx',
                        help='Output PowerPoint file name, default is "output/slides.pptx"')
    args = parser.parse_args()

    slides = parse_markdown(args.file)
    create_presentation(slides, args.output)

if __name__ == '__main__':
    main()
