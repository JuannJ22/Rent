from __future__ import annotations

from rentabilidad.gui import app as gui_app


def test_detects_remote_session_via_sessionname(monkeypatch):
    monkeypatch.setattr(gui_app, "_is_windows", lambda: True)
    monkeypatch.setenv("SESSIONNAME", "RDP-Tcp#3")
    monkeypatch.setenv("CLIENTNAME", "DESKTOP-REMOTE")

    assert gui_app._is_remote_session() is True


def test_local_console_not_flagged_as_remote(monkeypatch):
    monkeypatch.setattr(gui_app, "_is_windows", lambda: True)
    monkeypatch.setenv("SESSIONNAME", "Console")
    monkeypatch.setenv("CLIENTNAME", "Console")

    assert gui_app._is_remote_session() is False


def test_non_windows_environments_are_not_considered_remote(monkeypatch):
    monkeypatch.setattr(gui_app, "_is_windows", lambda: False)
    monkeypatch.delenv("SESSIONNAME", raising=False)
    monkeypatch.delenv("CLIENTNAME", raising=False)

    assert gui_app._is_remote_session() is False
