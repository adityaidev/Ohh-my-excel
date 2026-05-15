import sys
from pathlib import Path
from excel_graph_mcp.cli import main

FIXTURES = Path(__file__).parent / "fixtures"


def test_cli_build(monkeypatch, capsys):
    test_args = ["ohh-my-excel", "build", str(FIXTURES / "simple.xlsx")]
    monkeypatch.setattr(sys, "argv", test_args)
    main()
    captured = capsys.readouterr()
    assert "Built graph" in captured.out
    assert "nodes" in captured.out


def test_cli_status(monkeypatch, capsys):
    test_args = ["ohh-my-excel", "build", str(FIXTURES / "simple.xlsx")]
    monkeypatch.setattr(sys, "argv", test_args)
    main()
    test_args = ["ohh-my-excel", "status", str(FIXTURES / "simple.xlsx")]
    monkeypatch.setattr(sys, "argv", test_args)
    main()
    captured = capsys.readouterr()
    assert "Nodes" in captured.out


def test_cli_version(monkeypatch, capsys):
    test_args = ["ohh-my-excel", "version"]
    monkeypatch.setattr(sys, "argv", test_args)
    main()
    captured = capsys.readouterr()
    assert len(captured.out.strip()) > 0
