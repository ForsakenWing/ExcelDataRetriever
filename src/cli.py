import argparse


def parser() -> argparse.ArgumentParser.parse_args:
    _parser = argparse.ArgumentParser(description="Options to get specific path for output or input")
    template = _parser.add_argument_group(
        title="TEMPLATES",
        description="Use these arguments to create excel-template files"
    )
    template.add_argument(
        "--create-template",
        help="""Command to create excel template. By default create example file in templates folder.
        Can be overriden by prodiving path as arg after command e.g. template examples/
        Create excel sheet template in templates folder by default.
        Can be overwritten by providing path argument after template e.g.
        1. python main.py template templates/
        2. python main.py template
        3. python main.py template AbsPath/To/Template/Folder 
        """,
        dest="template",
        default=None,
        type=str,
        nargs="*"
    )
    _parser.add_argument(
        "--filepath", 
        help="Specify path in which program is looking for sheets. You can put relative or absolute path as well"
        "By default will look in sheets folder and current directory but not in sub-directories", 
        type=str,
        default="",
        metavar="Relative or absolute path e.g. sheets/ ; /Users/Test/folder "
    )
    _parser.add_argument(
        "--filename", 
        help="Casesensitive filename. Looks for file in sheets folder and current folder if filepath wasn't provided",
        type=str,
        default="",
        nargs=1,
        metavar="Filename e.g. IamSheet.xlsx"
    )
    _parser.add_argument(
        "--result_output",
        help="Path for result output",
        type=str,
        default="",
        nargs=1,
        metavar="Path/To/Output"
    )
    arguments = _parser.parse_args()
    return arguments

class Args:
    __args = parser()
    filepath = __args.filepath
    filename = __args.filename
    result_output = __args.result_output
    template = __args.template