from sys import argv
import logging
from docx import Document
from docx.table import _Cell as Cell_T
from docx.document import Document as Document_T
from pathlib import Path
from docx2pdf import convert
import win32api
from pywintypes import com_error
import pyzipper

DEBUG = False
OUT_PATH = Path("./out")

def convert_encrypt(files: dict[str, Path]):
    global OUT_PATH
    for birth_num, doc_path in files.items():
        try:
            pdf_path = OUT_PATH / f"{doc_path.stem}.pdf"
            zip_path = OUT_PATH / f"{doc_path.stem}.zip"

            convert(str(doc_path), str(pdf_path))
            with pyzipper.AESZipFile(
                    zip_path, 
                    'w',
                    compression=pyzipper.ZIP_LZMA,
                    encryption=pyzipper.WZ_AES) as z_out:
                z_out.setpassword(bytes(birth_num, 'UTF-8'))
                z_out.write(pdf_path, pdf_path.name)

            logging.debug(f'Created ZIP "{pdf_path}", password: {birth_num}')
            pdf_path.unlink()
        except com_error as err:
            logging.error(f'Docx2PDF error: "{win32api.FormatMessage(err.args[0])}"')
        except pyzipper.BadZipFile as err:
            logging.error(f'Zipping error')



def get_birth_number(doc_path: Path):
    doc = Document(str(doc_path))
    if not isinstance(doc, Document_T):
        logging.error(f'Failed opening "{doc_path}"')
    if len(doc.tables) <= 2:
        logging.error(f'No tables in "{doc_path}"')
    cell = doc.tables[2].cell(1,2)
    if isinstance(cell, Cell_T):
        return cell.text.strip()
    logging.error(f'Cell not found in "{doc_path}"')
    return None

def birth_number_gen(doc_paths: list[Path]):
    for doc_path in doc_paths:
        birth_num = get_birth_number(doc_path)
        yield (doc_path, birth_num)

def get_docx_files(dir_name: Path):
    logging.info(f'Starting on folder "{dir_name}"')
    return list(dir_name.glob('**/*.docx'))

#-----------------------
# MAIN FUNCTION AND INIT 
#-----------------------
def init_main():
    msg_format = "[%(levelname)s]: %(message)s"
    logging.basicConfig(level=logging.DEBUG, filename="log.txt", encoding="UTF-8", format=msg_format)
    logging.getLogger().name = "Podilnici"
    if not (OUT_PATH.exists() and OUT_PATH.is_dir()):
        OUT_PATH.mkdir(parents=True)
    if DEBUG:
        argv.append("./.local")
    if len(argv) == 1:
        logging.error("Not enough arguments, shutting down. Hint: Try dragging a file")
        exit(1)

def main():
    init_main()

    files = get_docx_files(Path(argv[1]))
    birth_nums = {
        birth_num: file
        for file, birth_num
        in birth_number_gen(files)
        if birth_num is not None
    }

    convert_encrypt(birth_nums)

if __name__=='__main__':
    main()