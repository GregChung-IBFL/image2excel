# image2excel
Converts an image file (.jpg, etc.) to a spreadsheet using individual cells colored red, green, or blue.  Cells are assigned numeric values (0-255) representing the RGB component value, and Excel conditional formatting rules set the cell colors.

- Coded in Python (3.9), with add-on libraries [OpenPyXL](https://openpyxl.readthedocs.io/), [Pillow](https://python-pillow.org/)
- Developed using Visual Studio Code; includes several sample launch configurations under (.vscode directory)

## Operation
image2excel can be easily run from the command line.  The only required argument is to pass in the name of the image file:
python image2excel Samples\ballons.jpg

Using a --preset option ("tiny", "small", "medium", "large") makes it easy to control the output size of the spreadsheet.

There are a number of other options as well; --help is available (via argparse).

When using Visual Studio Code, launch configurations (defined in .vscode\launch.json), it's easy to run the program with different settings.

## Sample Files
The Samples directory contains two sample images, balloons.jpg and sunflower.jpg.  The VSC launch configurations run the two files with various preset options.  In particular, sunflower.jpg is smaller than the target size of the large preset, and can be used to demonstrate the --enlarge option.

#### Version History
- v1: Initial version supporting command-line args to process a single file.  Includes four preset settings (tiny, small, medium, large).

## Special Thanks
The inspiration for this program came from Matt Parker's Stand Up Maths ["Stand-up comedy routine about Spreadsheets"](https://youtu.be/UBX2QQHlQ_I) video.  Beyond the idea, all project code was solely written by me.
