import ast
from pathlib import Path

import pandas as pd


def _load_prepare_studio_data():
    module_path = Path(__file__).resolve().parents[1] / "invoice_app_v3.py"
    source = module_path.read_text(encoding="utf-8")
    module_ast = ast.parse(source, filename=str(module_path))

    selected_nodes = [
        node
        for node in module_ast.body
        if isinstance(node, ast.FunctionDef)
        and node.name in {"prepare_studio_data"}
    ]

    compiled = compile(ast.Module(body=selected_nodes, type_ignores=[]), str(module_path), "exec")
    namespace = {"pd": pd}
    exec(compiled, namespace)
    return namespace["prepare_studio_data"]


prepare_studio_data = _load_prepare_studio_data()


def _base_row(project_ref, brief_ref, status, event="Event", description="Desc", owner="Owner"):
    return {
        'Project Ref': project_ref,
        'Event Name': event,
        'Project Description': description,
        'Project Owner': owner,
        'Brief Ref': brief_ref,
        'Content Brief Status': status,
    }


def test_prepare_studio_data_excludes_all_not_applicable_projects():
    df = pd.DataFrame([
        _base_row('SDG2000', 'BR1', 'Not Applicable'),
        _base_row('SDG2000', 'BR2', 'not applicable   '),
    ])

    result = prepare_studio_data(df)

    assert result.empty


def test_prepare_studio_data_sets_comment_for_non_completed_statuses():
    rows = [
        _base_row('SDG1000', 'BR1', 'Completed'),
        _base_row('SDG1000', 'BR2', 'Completed'),
        _base_row('SDG3000', 'BR3', 'Completed'),
        _base_row('SDG3000', 'BR4', 'In Progress'),
        _base_row('SDG4000', 'BR5', ''),
        _base_row('SDG4000', 'BR6', 'Not Applicable'),
    ]

    df = pd.DataFrame(rows)

    result = prepare_studio_data(df)
    result = result.set_index('Project Ref')

    assert set(result.index) == {'SDG1000', 'SDG3000', 'SDG4000'}

    # Projects with only completed lines should have no comment
    assert result.loc['SDG1000', 'Lines'] == 2
    assert result.loc['SDG1000', 'Studio Comment'] == ''

    # Projects with any non-completed status get the comment
    comment = 'check all lines are approved, artwork hours may require updating'
    assert result.loc['SDG3000', 'Lines'] == 2
    assert result.loc['SDG3000', 'Studio Comment'] == comment

    assert result.loc['SDG4000', 'Lines'] == 1
    assert result.loc['SDG4000', 'Studio Comment'] == comment
