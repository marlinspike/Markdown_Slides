import argparse
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
            # Remove the "Slide n:" part from the title
            title = line.split(': ', 1)[1] if ': ' in line else line
            current_slide = {'title': title, 'content': []}
        elif line.startswith('- '):  # Main bullet points
            current_slide['content'].append((line[2:], 'main'))
        elif line.startswith('  - '):  # Sub bullet points
            current_slide['content'].append((line[4:], 'sub'))

    if current_slide:
        slides.append(current_slide)
    return slides

def create_presentation(slides, pptx_file):
    """
    Creates a PowerPoint presentation from parsed slides.
    """
    prs = Presentation()
    bullet_slide_layout = prs.slide_layouts[1]  # 'Title and Content' layout
    
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

def main():
    parser = argparse.ArgumentParser(description="Convert Markdown to PowerPoint")
    parser.add_argument('-f', '--file', type=str, default='slides.md', help='Input markdown file name, default is "slides.md"')
    parser.add_argument('-o', '--output', type=str, default='slides.pptx', help='Output PowerPoint file name, default is "slides.pptx"')
    args = parser.parse_args()

    slides = parse_markdown(args.file)
    create_presentation(slides, args.output)

if __name__ == '__main__':
    main()
