import collections 
import collections.abc
from pptx import Presentation

file = './DiffusionModels_SB_v3.pptx'

ppt=Presentation(file)

notes = []

for page, slide in enumerate(ppt.slides):
    # this is the notes that doesn't appear on the ppt slide,
    # but really the 'presenter' note. 
    textNote = slide.notes_slide.notes_text_frame.text
    #notes.append((page,textNote))
    print("-------\nSlide {}: {}\n{}".format(page, slide.shapes.placeholders[1].text_frame.text, textNote))

#print(notes)
