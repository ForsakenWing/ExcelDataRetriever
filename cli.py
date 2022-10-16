from os import environ
from src import run_parser, create_template
from src.utils.excel_parser import write_excel_output

def main():
    create_template()
    write_excel_output(run_parser())

if __name__ == "__main__":
    main()