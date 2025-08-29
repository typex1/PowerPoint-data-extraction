# PowerPoint data extraction

## Motivation: extract all slide titles, slide content, and presenter notes and URLs from a PPTX file
For example, to have them available as a teleprompter text

Steps:

* Install relevant Python 3 module:
```
pip install python-pptx
```

* Run Python script:
```
python3 ./1-extract-presenter-notes.py <filename>.pptx
```
The output will be saved in **Mardown** format as <filename>_presenter_notes.md

Apart from slide data extraction, you can even create slides content - see more details here: https://python-pptx.readthedocs.io/en/latest/user/quickstart.html
