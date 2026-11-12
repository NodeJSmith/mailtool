"""stdout carries only JSON, so unattended consumers can always parse it.

Warnings, help and tracebacks go to stderr. The source check catches a bare
`print(...)` added later that would corrupt the JSON a caller parses.
"""

import ast
import json
import sys
from pathlib import Path

import pytest

from mailtool import cli

pytestmark = pytest.mark.unit

SOURCE_DIR = Path(cli.__file__).parent


def stdout_prints(path):
    """Yield (line, node) for each print() call in `path` that writes to stdout."""
    tree = ast.parse(path.read_text())
    for node in ast.walk(tree):
        if not (
            isinstance(node, ast.Call) and getattr(node.func, "id", None) == "print"
        ):
            continue
        if any(kw.arg == "file" for kw in node.keywords):
            continue
        yield node.lineno, node


def is_json_dumps(node):
    if len(node.args) != 1:
        return False
    arg = node.args[0]
    return (
        isinstance(arg, ast.Call)
        and isinstance(arg.func, ast.Attribute)
        and arg.func.attr == "dumps"
    )


@pytest.mark.parametrize("path", sorted(SOURCE_DIR.glob("*.py")), ids=lambda p: p.name)
def test_every_stdout_print_is_json(path):
    offenders = [
        f"{path.name}:{line}"
        for line, node in stdout_prints(path)
        if not is_json_dumps(node)
    ]
    assert offenders == [], (
        "print to stdout must be json.dumps(...); use file=sys.stderr"
    )


def run_main(monkeypatch, capsys, argv):
    monkeypatch.setattr(sys, "argv", ["mailtool", *argv])
    with pytest.raises(SystemExit) as exc:
        cli.main()
    captured = capsys.readouterr()
    return exc.value.code, captured.out, captured.err


def test_wrong_platform_prints_error_json(monkeypatch, capsys):
    monkeypatch.setattr(cli.sys, "platform", "linux")

    code, out, err = run_main(monkeypatch, capsys, ["emails"])

    assert code == 1
    assert json.loads(out)["status"] == "error"
    assert "requires Windows" in err


def test_missing_pywin32_prints_error_json(monkeypatch, capsys):
    monkeypatch.setattr(cli, "_check_platform", lambda: None)
    monkeypatch.setattr(cli.importlib.util, "find_spec", lambda _name: None)

    code, out, err = run_main(monkeypatch, capsys, ["emails"])

    assert code == 1
    assert json.loads(out) == {
        "status": "error",
        "message": "pywin32 is required but not installed",
    }
    assert "uv add pywin32" in err


def test_no_command_prints_help_on_stderr_and_error_json(monkeypatch, capsys):
    monkeypatch.setattr(cli, "_check_platform", lambda: None)
    monkeypatch.setattr(cli, "_check_pywin32", lambda: None)

    code, out, err = run_main(monkeypatch, capsys, [])

    assert code == 1
    assert json.loads(out) == {"status": "error", "message": "No command given"}
    assert "usage:" in err
