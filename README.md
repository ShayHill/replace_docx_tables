A test file to demonstrate how to insert paragraphs into a docx file with docx2python.

The workflow:

1. Create a template docx file with a table formatted the way you'd like
2. Insert a marker string into that table where you'd like your text to go
3. (see `test_assemble.py`): identify the paragraph, paragraph row, and
paragraph cell where the marker text is found
4. (see `test_assemble.py`): use these as templates to create new table cells
and rows with your input data
