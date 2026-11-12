import pytest


@pytest.fixture(autouse=True)
def _no_pc_state(monkeypatch):
    """Tests never read or write the per-PC last_run.json."""
    monkeypatch.setenv("MENU_CRAWLER_NO_STATE", "1")
