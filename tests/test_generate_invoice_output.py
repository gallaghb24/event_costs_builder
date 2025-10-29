import ast
from datetime import datetime
from pathlib import Path

import pandas as pd
from typing import Dict, Tuple


class _StubWorkbook:
    def __init__(self):
        self.sheetnames = []
        self.saved_paths = []
        self.closed = False

    def save(self, path):
        self.saved_paths.append(path)

    def close(self):
        self.closed = True


def _load_generate_invoice():
    module_path = Path(__file__).resolve().parents[1] / "invoice_app_v3.py"
    source = module_path.read_text(encoding="utf-8")
    module_ast = ast.parse(source, filename=str(module_path))

    selected_nodes = [
        node
        for node in module_ast.body
        if isinstance(node, ast.FunctionDef)
        and node.name in {"generate_invoice"}
    ]

    compiled = compile(ast.Module(body=selected_nodes, type_ignores=[]), str(module_path), "exec")

    namespace = {"pd": pd, "datetime": datetime, "Dict": Dict, "Tuple": Tuple}
    exec(compiled, namespace)
    return namespace["generate_invoice"]


def test_generate_invoice_uses_xlsm_extension_for_macro_templates(tmp_path):
    generate_invoice = _load_generate_invoice()

    stub_wb = _StubWorkbook()

    def _load_workbook(path, data_only=False, keep_vba=False):
        return stub_wb

    generate_invoice.__globals__["load_workbook"] = _load_workbook
    generate_invoice.__globals__["apply_formatting"] = lambda *_, **__: None

    template_info = {
        "path": str(tmp_path / "template.xlsm"),
        "formatting": {},
        "has_macros": True,
    }

    path, download_name, mime_type = generate_invoice(
        template_info,
        pd.DataFrame(),
        pd.DataFrame(),
        "Event Name",
        "E0000",
    )

    assert path.endswith(".xlsm")
    assert stub_wb.saved_paths == [path]
    assert download_name.endswith(".xlsm")
    assert mime_type == "application/vnd.ms-excel.sheet.macroEnabled.12"


def test_generate_invoice_uses_xlsx_extension_when_no_macros(tmp_path):
    generate_invoice = _load_generate_invoice()

    stub_wb = _StubWorkbook()

    def _load_workbook(path, data_only=False, keep_vba=False):
        return stub_wb

    generate_invoice.__globals__["load_workbook"] = _load_workbook
    generate_invoice.__globals__["apply_formatting"] = lambda *_, **__: None

    template_info = {
        "path": str(tmp_path / "template.xlsx"),
        "formatting": {},
        "has_macros": False,
    }

    path, download_name, mime_type = generate_invoice(
        template_info,
        pd.DataFrame(),
        pd.DataFrame(),
        "Event Name",
        "E0000",
    )

    assert path.endswith(".xlsx")
    assert stub_wb.saved_paths == [path]
    assert download_name.endswith(".xlsx")
    assert mime_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def test_generate_invoice_adds_comments_to_studio_sheet(tmp_path):
    generate_invoice = _load_generate_invoice()

    class _Cell:
        def __init__(self):
            self.value = None
            self.comment = None

    def _column_letter(idx):
        letters = ''
        while idx > 0:
            idx, remainder = divmod(idx - 1, 26)
            letters = chr(65 + remainder) + letters
        return letters

    class _Sheet:
        def __init__(self):
            self.cells = {}
            self.max_row = 10
            self.max_column = 10

        def _ensure_cell(self, coord):
            if coord not in self.cells:
                self.cells[coord] = _Cell()
            return self.cells[coord]

        def __getitem__(self, coord):
            return self._ensure_cell(coord)

        def __setitem__(self, coord, value):
            self._ensure_cell(coord).value = value

        def cell(self, row, column):
            coord = f"{_column_letter(column)}{row}"
            return self._ensure_cell(coord)

    class _WorkbookWithStudio:
        def __init__(self):
            self.sheetnames = ['Studio']
            self.studio = _Sheet()
            self.saved_paths = []
            self.closed = False

        def __getitem__(self, name):
            if name == 'Studio':
                return self.studio
            raise KeyError(name)

        def save(self, path):
            self.saved_paths.append(path)

        def close(self):
            self.closed = True

    stub_wb = _WorkbookWithStudio()

    def _load_workbook(path, data_only=False, keep_vba=False):
        return stub_wb

    generate_invoice.__globals__["load_workbook"] = _load_workbook
    generate_invoice.__globals__["apply_formatting"] = lambda *_, **__: None
    generate_invoice.__globals__["Comment"] = lambda text, author: {"text": text, "author": author}

    template_info = {
        "path": str(tmp_path / "template.xlsx"),
        "formatting": {"Studio": {}},
        "has_macros": False,
    }

    studio_df = pd.DataFrame([
        {
            "Project Ref": "SDG1000",
            "Event Name": "Event",
            "Project Description": "Desc",
            "Project Owner": "Owner",
            "Lines": 2,
            "Studio Hours": None,
            "Type": "Artwork",
            "Core/OAB": "CORE",
            "Studio Comment": "Needs review",
        }
    ])

    generate_invoice(
        template_info,
        studio_df,
        pd.DataFrame(),
        "Event",
        "E0000",
    )

    assert stub_wb.studio.cells["A3"].comment == {"text": "Needs review", "author": "Status"}
