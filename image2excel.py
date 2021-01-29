"""
    image2excel : Creates a spreadsheet containing a cells-as-pixels representation of a source
    image.  Each column in the output spreadsheet represents one of the primary colors (Red, Green,
    Blue).  Cells are assigned numerical values between 0-255 representing the brightness of that
    column's color.  Thus, three adjacent cells in a row represent one RGB pixel.  For example,
    HTML LightCoral is #F08080, corresponding to decimal RGB values of (240,128,128).  A
    LightCoral image pixel would be represented in cells A1:A3 with values of 240, 128, and 128.
    When zoomed in, you can see the individual cells' red, green, and blue component colors.  When
    zoomed out and/or standing far back, light combines to create the color palette we expect.
    Thus we can see the whole image, albeit possibly a little darker and grainier.

    Coded by Greg Chung : https://github.com/GregChung-IBFL/image2excel

    V1.0    Tool supports command-line args, presets for default output sizes, operates on a single file.
    V1.01   Bug fix for GIF support.
    V1.1    Adds support for batch processing.
"""

from os import path
from datetime import datetime
import sys
import glob
import json
import argparse
import textwrap

import image_converter  # Mine

# Various application default settings, including the definitions of the presets.
# Many values are overridden via the command-line args, if specified.
CONFIG_FILE_APP = "config.json"


class Application :
    """Main Application class for image2excel.  Handles the housekeeping, while using the image_converter
    module to do the actual processing one an image.
    """

    def __init__(self) :
        self.config = {}


    def read_config_file(self) :
        """Loads basic configuration from the config file into the self.config dict.
        It is a fatal error if the config file cannot be loaded.
        """
        try:
            with open( CONFIG_FILE_APP, "r") as jsonfile:
                self.config = json.load( jsonfile )
        except:
            print( F'Failed to read application configuration file "{CONFIG_FILE_APP}", aborting!' )
            sys.exit( -990 )


    def parse_command_line(self) :
        """Parses the command line arguments, merging them into the config dict.
        input_file is required.  All other settings are optional and will revert to
        default values, many themselves being assigned defaults via the config file.
        """
        epilog = textwrap.dedent( """
        Batch Mode:
            Batch mode is automatically enabled if <input_file> refers to a directory, instead of an individual
            file.  In batch mode, all image files found in the named directory are processed.  Batching is
            shallow; images located in subdirectories under the batch directory are not processed.
            If <output_file> is specified with batch mode, it is required to be a directory to receive the
            generated files.  If <output_file> is not specified, the batch processed output files will be
            saved to the same directory as <input_file>.
        Presets:
            small =  """ + str(self.config["presets"]["small"]) + """
            medium = """ + str(self.config["presets"]["medium"]) + """
            large =  """ + str(self.config["presets"]["large"]) + """
        Use of presets will override --output_height and --output_width, if specified.
        """ )

        parser = argparse.ArgumentParser( description = "Converts an image file to a spreadsheet represention using conditional formatting rules",
                                            formatter_class = argparse.RawDescriptionHelpFormatter,
                                            epilog = epilog
                                        )
        parser.add_argument( "input_file", help = "Input image filename/directory" )
        parser.add_argument( "output_file", default = "", nargs = "?", help = "Output filename/directory, defaults to <input_file>.xlsx if not specified" )
        parser.add_argument( "--output_zoom", type = float, default = self.config["output_zoom"], help = "Desired zoom ratio" )
        parser.add_argument( "--output_col_width", type = float, default = self.config["output_col_width"], help = "Desired output column width" )
        parser.add_argument( "--output_row_height", type = float, default = self.config["output_row_height"], help = "Desired output row height" )
        parser.add_argument( "--output_height", type = float, default = self.config["output_height"], help = "Image output height (px)" )
        parser.add_argument( "--output_width", type = float, default = self.config["output_width"], help = "Image output width (px)" )
        parser.add_argument( "--enlarge", action = argparse.BooleanOptionalAction, help = "If enabled, smaller images can be enlarged; otherwise, images can only be reduced" )
        parser.add_argument( "--preset", choices = ["tiny", "small", "medium", "large"], help = "Use preset settings (see below)" )

        # Parse the command line.  args isa argparse.Namespace, vars(args) returns a dict.
        args = parser.parse_args()
        args_dict = vars(args)

        # Apply preset settings, if one is specified:
        preset = args_dict["preset"]
        if preset :
            print( F'Using "{preset}" preset.' )
            args_dict.update( self.config["presets"][preset] )

        # Merge into the config dict:
        self.config.update( args_dict )



    def run(self) :
        """Runs the app!"""
        print("")
        print("image2excel: Converts image file into a spreadsheet of individual Red, Green, Blue cells.")
        print("")

        self.read_config_file()

        self.parse_command_line()

        # Test if input_file exists and is a directory --> enable batch mode!
        input_file = self.config["input_file"]
        is_file = path.isfile( input_file )
        is_dir = path.isdir( input_file )

        # If input_file is an individual file, put its name into our files_list list.  If it
        # refers to a directory, glob up all the file names into files_list.
        self.config["batch_mode"] = False
        if is_dir :
            # Batch mode!
            self.config["batch_mode"] = True

            # If the output_file is named, it must refer to an existing directory.
            if not path.isdir( self.config["output_file"] ) :
                print( F'Batch output target "{self.config["output_file"]}" does not exist or is not a directory.' )
                sys.exit( -1 )

            print( F'Batch processing image files in "{input_file}".' )
            files_list = glob.glob( input_file + "\\*" )
        elif is_file :
            files_list = [ input_file ]
        else :
            print( F'Input file "{input_file}" does not exist, verify path is valid.' )
            sys.exit( -1 )

        start_time = datetime.now()
        num_files_seen = 0
        num_files_processed = 0

        for file_path in files_list :
            # In batch mode, files_list will contain everything directly under the named
            # directory, which can include subdirectories.  This tool does not recurse,
            # so skip anything which isn't a file.
            if not path.isfile( file_path ) :
                continue

            # Pillow docs recommend simply trying to open all files using Image.open()
            # to find the image files.  i.e. don't need to pre-filter by file extension.
            # https://pillow.readthedocs.io/en/stable/handbook/tutorial.html

            num_files_seen += 1
            converter = image_converter.Converter( file_path, self.config )
            if converter.process_file() :
                num_files_processed += 1

        # Calc the total processing time, then report some stats
        elapsed = (datetime.now() - start_time).total_seconds()
        elapsed_str = []
        elapsed_minutes = elapsed // 60
        if elapsed_minutes > 0 :
            elapsed_str.append( F"{int(elapsed_minutes)} minutes" )
        elapsed_seconds = elapsed % 60
        if elapsed_seconds >= 0 :
            elapsed_str.append( F"{int(elapsed_seconds)} seconds" )

        print("")
        print( F"Files seen: {num_files_seen}" )
        print( F"Files processed: {num_files_processed}" )
        print(  "Elapsed time: " + ", ".join(elapsed_str) )


    @staticmethod
    def main() :
        """Instantiate the Application class and run."""
        app = Application()
        app.run()




if __name__ == "__main__" :
    Application.main()
