import collections 
import collections.abc
from pptx import Presentation
import re

file = './Module5.pptx'
ppt=Presentation(file)
notes = []

for page, slide in enumerate(ppt.slides):
    # this is the notes that doesn't appear on the ppt slide,
    # but really the 'presenter' note. 
    text_notes = slide.notes_slide.notes_text_frame.text
    try:
        # Regular expression pattern to match HTTPS links
        pattern = r'https://\S+'
        #print("TST: Slide {}, length {}".format(int(page)+1, len(slide.shapes.placeholders)))
        text_frame_text = slide.shapes.placeholders[1].text_frame.text
    except Exception as error:
        text_frame_text = slide.shapes.placeholders[0].text_frame.text
        
    print("\n-------\n### Slide {}: {}\n\n{}".format(int(page)+1, text_frame_text, text_notes))
    # Use re.findall() to extract all matches
    https_links = re.findall(pattern, text_notes)
    if len(https_links) > 0:
        print("\nHTTPS links found:")
        for link in https_links:
            cleaned_link = re.sub(r'[^/\w\d]+$', '', link)
            print("* {}".format(cleaned_link))
