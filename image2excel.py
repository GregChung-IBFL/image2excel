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
"""

import json
import argparse

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
            exit( -990 )


    def parse_command_line(self) :
        """Parses the command line arguments, merging them into the config dict.
        input_file is required.  All other settings are optional and will revert to
        default values, many themselves being assigned defaults via the config file.
        """

        epilog = "Presets:\n"
        epilog += "    small =  " + str(self.config["presets"]["small"]) + "\n"
        epilog += "    medium = " + str(self.config["presets"]["medium"]) + "\n"
        epilog += "    large =  " + str(self.config["presets"]["large"]) + "\n"
        epilog += "Use of presets will override --output_height and --output_width, if specified."

        parser = argparse.ArgumentParser( description = "Converts an image file to a spreadsheet represention using conditional formatting rules",
                                            formatter_class = argparse.RawDescriptionHelpFormatter,
                                            epilog = epilog
                                        )
        parser.add_argument( "input_file", help = "Input image filename" )
        parser.add_argument( "output_file", default = "", nargs = "?", help = "Output filename, defaults to <input_file>.xlsx if not specified" )
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

        filepath = self.config["input_file"]
        converter = image_converter.Converter( filepath, self.config )

        converter.process_file()

        print("\nDone!")





    @staticmethod
    def main() :
        """Instantiate the Application class and run."""
        app = Application()
        app.run()




if __name__ == "__main__" :
    Application.main()
