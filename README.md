# image2excel
Converts an image file (.jpg, etc.) to a spreadsheet using individual cells colored red, green, or blue.  Cells are assigned numeric values (0-255) representing the RGB component value, and Excel conditional formatting rules set the cell colors.

- Coded in Python (3.9), with add-on libraries [OpenPyXL](https://openpyxl.readthedocs.io/), [Pillow](https://python-pillow.org/)
- Developed using Visual Studio Code; includes several sample launch configurations under (.vscode directory)
- Supports batch processing (version 1.1)

![Screenshot](https://user-images.githubusercontent.com/76141087/105903846-64a01500-5fee-11eb-90f9-b4a6bf79ef8e.png)

## What it Does / How the Spreadsheet Works
The program converts the image into a spreadsheet containing the pixel-by-pixel Red, Green, and Blue component colors.  Each column in the spreadsheet represents a single primary color (R,G,B).  A cell contains a numeric value between 0 and 255, reflecting the RGB value making up the image pixel.  Three horizontally adjacent cells, such as A1:C1, thus contain the amount of Red, Green, and Blue in one pixel.

![RGB numeric values](https://user-images.githubusercontent.com/76141087/105903481-ecd1ea80-5fed-11eb-98a0-4e3e7fd12cd6.png)

Conditional Formatting rules are applied to every column.  The rules set the cell to a solid color using a two color gradient scale, from black (at value 0) to the appropriate RGB color (at value 255).  Every third column (A, D, G, ...) is red, the adjacent columns (B, E, H, ...) are green, and the remaining third columns (C, F, I, ...) are blue.  Increasing the magnification (the little slider in the lower-right in Excel), you should be able to see how the columns are discrete colors.

![Magnified discrete RGB coloration](https://user-images.githubusercontent.com/76141087/105902200-4fc28200-5fec-11eb-9a9b-306f5f319f1b.png)

The Conditional Formatting rules themselves can be viewed inside Excel.  From the Home menu select Conditional Formatting > Manage Rules (alternatively, use keyboard accelerators Alt-H > L > R).  Change the "Show formatting rules for" option to "This Worksheet" and you will see the three Graded Color Scale rules are applied to all the cells representing the image.

![Conditional Formatting](https://user-images.githubusercontent.com/76141087/105902205-50f3af00-5fec-11eb-9d54-a883672bb5b2.png)

### Dependencies
The project contains a requirements.txt file.  Dependencies can be installed using pip: <br/>
`pip install -r requirements.txt`

### Operation
image2excel can be easily run from the command line.  The only required argument is to pass in the name of the image file.  The following will convert the sample image using default settings, creating balloons.xlsx in the same directory: <br/>
`python image2excel.py Samples\balloons.jpg`

Use a `--preset` option ("tiny", "small", "medium", "large") to quickly control the output size of the spreadsheet: <br/>
`python image2excel.py Samples\balloons.jpg --preset large`

When not using a preset, the output height and/or width resolution can be specified with `--output_height` and `--output_width`: <br/>
`python image2excel.py Samples\balloons.jpg --output_height 100`

The `--enlarge` option allows the tool to upsize small source images into large spreadsheets: <br/>
`python image2excel.py Samples\sunflower.jpg --enlarge --preset large`

`--output_zoom` controls the zoom percentage when opened.  The following will cause the spreadsheet to be opened at 200% magnification: <br/>
`python image2excel.py Samples\balloons.jpg --output_zoom 200`

`--help` is available!

When using Visual Studio Code, launch configurations (defined in .vscode\launch.json), it's easy to run the program with different settings.

### Batch Processing
image2excel can process a batch of multiple image files, when the named input file refers to a directory instead of a specific file.  All image files found under the input directory will be converted to spreadsheets.  By default, the generated .xlsx files will be placed in the same directory as the input images, but an alternative directory to receive the files can be specified.  The following command will process all images found within the Samples directory, and place the spreadsheet files into a directory named Output: <br />
`python image2excel.py Samples Output`

Settings, such as using a `--preset`, are applied to all files within the batch.

### Image Resizing
Before being converted to a spreadsheet, the source images are resized to fit the desired dimensions.  The `--output_height` and `--output_width` values determine the target image size.  If a `--preset` is named, the preset values will override any values specified by `--output_height` and `--output_width`.
If the `--enlarge` option is specified, small images can be resized to expand into the requested sizes, if necessary.  Without this option, images will only reduce (shrink).  If a small image is used with larger requested dimensions, but `--enlarge` is not specified, the spreadsheet image will use the source's original dimensions, not the requested dimensions.
During resizing, the aspect ratio is maintained.


#### Sample Files
The Samples directory contains two sample images, balloons.jpg and sunflower.jpg.  The VSC launch configurations run the two files with various `--preset` options.  In particular, sunflower.jpg is a small image which can be used to demonstrate the `--enlarge` option.


#### Version History
- v1: Initial version supporting command-line args to process a single file.  Includes four preset settings (tiny, small, medium, large).
- v1.01: Fixes support for GIF format (palette mode instead of RGB mode).
- v1.1: Adds batch support to process all image files under the named directory.

### Special Thanks
The inspiration for this program came from Matt Parker's Stand-up Maths ["Stand-up comedy routine about Spreadsheets"](https://youtu.be/UBX2QQHlQ_I) video.  Beyond the idea, all project code was solely written by me.
