"""Shared test fixtures."""

import pytest


@pytest.fixture(autouse=True)
def isolated_icons_dir(tmp_path, monkeypatch):
    """Point the local custom-icon library at an empty per-test folder so the
    suite never picks up icons that happen to live in the developer's home
    directory. Tests that exercise local icons write files into the returned
    path."""
    icons = tmp_path / "icons"
    icons.mkdir()
    monkeypatch.setenv("VISIO_MCP_ICONS_DIR", str(icons))
    return icons
