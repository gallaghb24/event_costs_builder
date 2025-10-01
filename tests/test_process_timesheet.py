import ast
import io
import math
import os
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd


class _StreamlitStub:
    def __init__(self):
        self.errors = []

    def error(self, message):
        self.errors.append(message)


class _UploadedFile(io.BytesIO):
    def getbuffer(self):  # pragma: no cover - exercised indirectly
        return super().getbuffer()


def _load_process_timesheet():
    module_path = Path(__file__).resolve().parents[1] / "invoice_app_v3.py"
    source = module_path.read_text(encoding="utf-8")
    module_ast = ast.parse(source, filename=str(module_path))

    selected_nodes = [
        node
        for node in module_ast.body
        if isinstance(node, ast.FunctionDef)
        and node.name in {"round_up_to_quarter", "process_timesheet"}
    ]

    compiled = compile(ast.Module(body=selected_nodes, type_ignores=[]), str(module_path), "exec")

    st_stub = _StreamlitStub()
    namespace = {
        "pd": pd,
        "np": np,
        "tempfile": tempfile,
        "os": os,
        "math": math,
        "st": st_stub,
    }

    exec(compiled, namespace)
    return namespace["process_timesheet"], st_stub


def _build_sample_csv() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Job Number": ["1/SDG1234", "1/SDG1234"],
            "Job Description": ["Test Job", "Test Job"],
            "Charge Code": ["Artwork", "Artwork"],
            "Total": [1.0, 1.0],
        }
    )


def test_process_timesheet_reads_utf8_csv(tmp_path):
    process_timesheet, st_stub = _load_process_timesheet()
    csv_bytes = _build_sample_csv().to_csv(index=False).encode("utf-8")
    uploaded = _UploadedFile(csv_bytes)

    result = process_timesheet(uploaded)

    expected = pd.DataFrame(
        {
            "Project Ref": ["SDG1234"],
            "Total Hours": [2.0],
            "Type": ["Artwork"],
            "Core/OAB": ["CORE"],
        }
    )

    pd.testing.assert_frame_equal(result.reset_index(drop=True), expected)
    assert not st_stub.errors


def test_process_timesheet_falls_back_to_utf16(tmp_path):
    process_timesheet, st_stub = _load_process_timesheet()
    csv_bytes = _build_sample_csv().to_csv(index=False).encode("utf-16")
    uploaded = _UploadedFile(csv_bytes)

    result = process_timesheet(uploaded)

    expected = pd.DataFrame(
        {
            "Project Ref": ["SDG1234"],
            "Total Hours": [2.0],
            "Type": ["Artwork"],
            "Core/OAB": ["CORE"],
        }
    )

    pd.testing.assert_frame_equal(result.reset_index(drop=True), expected)
    assert not st_stub.errors
