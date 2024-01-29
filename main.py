from sys import argv
import logging
from docx import Document
from docx.table import _Cell as Cell_T
from docx.document import Document as Document_T
from glob import glob
from pathlib import Path

DEBUG = True
OUT_PATH = Path("./out")

def get_birth_number(doc_path):
    doc = Document(doc_path)
    if not isinstance(doc, Document_T):
        logging.error(f'Failed opening "{doc_path}"')
    if len(doc.tables) <= 2:
        logging.error(f'No tables in "{doc_path}"')
    cell = doc.tables[2].cell(1,2)
    if isinstance(cell, Cell_T):
        return cell.text
    logging.error(f'Cell not found in "{doc_path}"')
    return None

def get_docx_files(dir_name: Path):
    logging.info(f'Starting on folder "{dir_name}"')
    files = glob(str(dir_name / "*/*.docx"), recursive=True)
    files.extend(glob(str(dir_name / "*.docx"), recursive=True))
    return files

#-----------------------
# MAIN FUNCTION AND INIT 
#-----------------------
def init_main():
    logging.basicConfig(level=logging.INFO, filename="log.txt")
    logging.getLogger().name = "Podilnici"
    if DEBUG:
        argv.append("./.local")

def main():
    init_main()
    if len(argv) == 1:
        logging.error("Not enough arguments, shutting down. Hint: Try dragging a file")
        exit(1)
    files = get_docx_files(Path(argv[1]))
    birth_nums = {file: get_birth_number(file) for file in files}
    print(birth_nums)


if __name__=='__main__':
    main()