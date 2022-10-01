from openpyxl import load_workbook, Workbook
from openpyxl.utils.exceptions import InvalidFileException
from .cli import Args
import os
import logging

logging.basicConfig(format="%(levelname)s - %(message)s", level=logging.ERROR)

def list_to_str_serializer(lst: list[str]) -> str:
    try:
        tmp = "".join(lst)
    except (ValueError, AttributeError):
        tmp = ""
        logging.info("Serializing arguments...")
    return tmp
        
def filename_handler(filename: str = list_to_str_serializer(Args.filename)) -> load_workbook:
    if not check_right_format(filename):
        filename = f"{filename}.xlsx"
    try:
        workbook = load_workbook(filename)
    except FileNotFoundError:
        logging.error("Whoops..Can't find this file")
    except InvalidFileException:
        logging.error("Whoops..Supported formats are: .xlsx,.xlsm,xltx,.xltm")
    else:
        return workbook    

    
def check_right_format(filename: str) -> bool:
    return filename.endswith((".xlsx", ".xlsm", ".xltx", ".xltm"))


def collect_excel_files(filepath: str) -> list[str] | None:
    files_result = []
    for (root, _, files) in os.walk(filepath):
        for file in files:
             if check_right_format(file): 
                 files_result.append(os.path.join(root, file))
    return files_result
                 
def filepath_handler(filepath: str = list_to_str_serializer(Args.filepath)) -> list[str] | None | str:
    if os.path.isdir(filepath):
        if os.path.isabs(filepath):
            files_result = collect_excel_files(filepath)
            return files_result
        else:
            files_result = collect_excel_files(os.path.join(os.getcwd(), filepath))
            return files_result
    if os.path.isfile(filepath):
        return filename_handler(filepath)
    else:
        filepath = os.path.join(os.getcwd(), "sheets")
        files_result = collect_excel_files(filepath)
        return files_result + [os.path.join(os.getcwd(), file) for file in os.listdir() if check_right_format(file)]


def check_and_update_file_format(filename):
    if not check_right_format(filename):
        return f"{filename}.xlsx"


def create_excel_file(dirname: str, filename: str = "example.xlsx"):
    if not filename:
        filename="example.xlsx"
    if (_file_name := check_and_update_file_format(filename)) is not None: filename = _file_name
    try:
        with open(os.path.join(dirname, filename), "x") as file:
                    file.write("MOCK")
    except FileExistsError:
        with open(os.path.join(dirname, filename), "w") as file:
            file.write("MOCK")


def create_template(path: str = Args.template) -> None:
    if path is None:
        return
    if not path:
        dirname= "templates"
        try:
            os.mkdir(dirname)
        except OSError:
            logging.info("templates directory already exists")
        finally:
            create_excel_file(dirname)
    else:
        head, tail = os.path.split("".join(path))
        if not head:
            head, tail = tail, None
        os.makedirs(head, exist_ok=True)
        create_excel_file(head, tail)
        
def workbook_serializer(wb: Workbook):
    worksheet = wb.active
    cell_identifiers = [(cell.row, cell.column, cell.value) for cell in worksheet['1'] if cell.value is not None]
    result = []
    table_index = 1
    while True:
        tmp = dict()
        for row_index, column_index, key in cell_identifiers:
            cell = worksheet.cell(row_index + table_index, column_index)
            tmp[key] = cell.value
        if len(set(tmp.values())) <= 1:
            return result
        else:
            table_index += 1
            result.append(tmp)

def run_converter() -> list[dict]:
    if Args.filename and (workbook := filename_handler()):
        return workbook_serializer(workbook)
    excel_files = filepath_handler()
    if isinstance(excel_files, Workbook):
        return workbook_serializer(excel_files)
    worksheets: list[Workbook] = [filename_handler(wb) for wb in excel_files]
    user_data = sum([workbook_serializer(worksheet) for worksheet in worksheets], []) # Removing arrays nesting
    return user_data