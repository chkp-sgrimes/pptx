Script that pulls text from powerpoint slides

**Flow**
- open pptx files in inputs directory
- try to open each pptx
    - pptx does not open, move it to the encrypted directory
    - pptx opens, process each slide
- slides have shapes; line, picture, text frame, groups, etc
- example slide hierarchy
    - text frame and groups
        - group with text frames
            - group with text frames
                - ...
- recurse through the hierarchy of groups in the slide to retrieve text
- write the text to outputs/[powerpoint name]_[slide number].txt
