Script that pulls text from powerpoint slides using the python-pptx module

Powerpoint slides have shapes; line, picture, text frame, groups, etc
- example slide hierarchy
    - text frame and groups
        - group with text frames
            - group with text frames
                - ...
**Flow**
- open pptx files in inputs directory
- try to open each pptx
    - pptx does not open, move it to the encrypted directory
    - pptx opens, process each slide
- recurse through the hierarchy of groups in the slide to retrieve text
- write the text to outputs/[powerpoint name]_[slide number].txt
