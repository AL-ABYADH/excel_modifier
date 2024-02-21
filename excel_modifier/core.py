import os
import zipfile
from io import BytesIO
from bs4 import BeautifulSoup


class Excel:
    def __init__(self, file_path):
        self.path = file_path
        self._validate_file_path()
        self.active_sheet = 1
        self._sheet_content, self._sheet_info = self._get_sheet_content_and_info()

    def _validate_file_path(self):
        if not os.path.isfile(self.path):
            raise FileNotFoundError(f"The file {self.path} was not found!")

    def get_worksheets(self):
        with zipfile.ZipFile(self.path) as zip_ref:
            soup = BeautifulSoup(zip_ref.read("xl/workbook.xml"), features="xml")
            return [sheet.get("name") for sheet in soup.find_all("sheet")]

    def _get_sheet_content_and_info(self):
        with zipfile.ZipFile(self.path) as zip_ref:
            file_path = f"xl/worksheets/sheet{self.active_sheet}.xml"
            content = BeautifulSoup(zip_ref.read(file_path), features="xml")
            sheet_info = zip_ref.getinfo(file_path)
        return content, sheet_info

    def change_value(self, cell, value):
        cell_tag = self._sheet_content.find("c", attrs={"r": cell})
        if cell_tag and (v_tag := cell_tag.find("v")):
            v_tag.string = str(value)
        else:
            print(f"Cell {cell} not found. Consider implementing cell creation logic here.")

    def set_active_sheet(self, sheet):
        self.save()  # Save current sheet changes before switching
        worksheets = self.get_worksheets()
        if isinstance(sheet, int):
            self.active_sheet = sheet
        elif isinstance(sheet, str) and sheet in worksheets:
            self.active_sheet = worksheets.index(sheet) + 1
        else:
            raise ValueError("Sheet identifier not recognized or not found.")
        self._sheet_content, self._sheet_info = self._get_sheet_content_and_info()

    def save(self):
        with BytesIO() as temp_zip_contents, zipfile.ZipFile(self.path) as zip_ref:
            with zipfile.ZipFile(temp_zip_contents, 'w', zipfile.ZIP_DEFLATED) as temp_zip:
                for item in zip_ref.infolist():
                    if item.filename != f"xl/worksheets/sheet{self.active_sheet}.xml":
                        temp_zip.writestr(item, zip_ref.read(item.filename))
                temp_zip.writestr(self._sheet_info, str(self._sheet_content))

            with open(self.path, 'wb') as file:
                file.write(temp_zip_contents.getvalue())