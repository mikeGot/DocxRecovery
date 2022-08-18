import zipfile
from pathlib import Path
from customExceptions import BadZipFileException, ChangeFileExtensionException, \
    RecoverIsNotPossible, ParseXmlException, SaveFileException

from bs4 import BeautifulSoup as bs


def change_file_extension(filename: Path, new_extension: str) -> Path:
    """
    :param filename:
    :param new_extension: example - .zip or .docx
    :return Path:
    """
    try:
        f = filename.rename(filename.with_suffix(new_extension))
        return f
    except Exception as e:
        print(filename, e)
        raise ChangeFileExtensionException


# change_file_extension(Path("/Users/mike/Desktop/document.docx"), ".zip")


def file_unarchive(filename: Path) -> Path:
    try:
        directory_to_extract = Path(filename.with_suffix(''))
        with zipfile.ZipFile(filename, 'r') as zip_ref:
            zip_ref.extractall(path=directory_to_extract)
        return directory_to_extract
    except zipfile.BadZipfile as bz:
        print(filename, bz)
        raise BadZipFileException


def path_to_xml(path_to_dir: Path) -> Path:
    special_file_for_recover: str = "word/document.xml"
    return path_to_dir.joinpath(special_file_for_recover)


def if_recover_possible(path_to_file: Path) -> bool:
    if path_to_file.exists():
        return True
    else:
        raise RecoverIsNotPossible


def parse_xml(xml_file: Path) -> str:
    try:
        content = []
        # Read the XML file
        with open(xml_file, "r") as file:
            # Read each line in the file, readlines() returns a list of lines
            content = file.readlines()
            # Combine the lines in the list into a string
            content = "".join(content)
            bs_content = bs(content, "lxml")

            result = bs_content.findAll("w:t")
        return " ".join([r.text for r in result])
    except Exception as e:
        print(xml_file, e)
        raise ParseXmlException


def save_file(result_str: str, file: Path):
    try:
        with open(file, "w+") as f:
            f.write(result_str)
    except Exception as e:
        print(file, e)
        raise SaveFileException


if __name__ == "__main__":
    path_to_docx = Path("/Users/mike/Desktop/dop.docx")
    path_to_zip = change_file_extension(path_to_docx, ".zip")
    path_to_zip_dir = file_unarchive(path_to_zip)
    change_file_extension(path_to_zip, ".docx")
    path_to_xml_file = path_to_xml(path_to_zip_dir)
    if if_recover_possible(path_to_xml_file):
        result_text = parse_xml(path_to_xml_file)
        save_file(result_text, path_to_zip_dir.with_suffix(".txt"))

