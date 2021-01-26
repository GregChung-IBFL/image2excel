# image2excel
Converts an image file (.jpg, etc.) to a spreadsheet using individual cells colored red, green, or blue.  Cells are assigned numeric values (0-255) representing the RGB component value, and Excel conditional formatting rules set the cell colors.

- Coded in Python (3.9), with add-on libraries [OpenPyXL](https://openpyxl.readthedocs.io/), [Pillow](https://python-pillow.org/)
- Developed using Visual Studio Code; includes several sample launch configurations under (.vscode directory)

## Special Thanks
The inspiration for this program came from Matt Parker's Stand Up Maths ["Stand-up comedy routine about Spreadsheets"](https://youtu.be/UBX2QQHlQ_I) video.
However, all code was written by me.

#### Version History
- v1: Initial version supporting command-line args to process a single file.  Includes four preset settings (tiny, small, medium, large).
