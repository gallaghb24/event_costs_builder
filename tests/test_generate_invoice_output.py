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
