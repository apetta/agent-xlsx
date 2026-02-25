"""Tests for engine resolution: AGENT_XLSX_ENGINE env var override and Aspose detection."""

import pytest

from agent_xlsx.utils.engine import resolve_engine


def test_env_var_overrides_auto(monkeypatch):
    """AGENT_XLSX_ENGINE env var overrides 'auto' engine selection."""
    # Mock Aspose as unavailable so auto would normally fall through
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter._ASPOSE_AVAILABLE", None)
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter.is_aspose_available", lambda: False)
    monkeypatch.setattr("agent_xlsx.adapters.xlwings_adapter.is_excel_available", lambda: False)

    # Set env var to libreoffice and mock it as available
    monkeypatch.setenv("AGENT_XLSX_ENGINE", "libreoffice")
    monkeypatch.setattr(
        "agent_xlsx.adapters.libreoffice_adapter.is_libreoffice_available",
        lambda: True,
    )

    result = resolve_engine("screenshot", "auto", libreoffice=True)
    assert result == "libreoffice"


def test_env_var_ignored_when_explicit_engine(monkeypatch):
    """AGENT_XLSX_ENGINE is ignored when user explicitly sets engine != auto."""
    monkeypatch.setenv("AGENT_XLSX_ENGINE", "libreoffice")
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter.is_aspose_available", lambda: True)

    # Explicit "aspose" should not be overridden by the env var
    result = resolve_engine("screenshot", "aspose", libreoffice=True)
    assert result == "aspose"


def test_aspose_unavailable_falls_through(monkeypatch):
    """When Aspose is unavailable, auto falls through to next engine."""
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter._ASPOSE_AVAILABLE", None)
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter.is_aspose_available", lambda: False)
    monkeypatch.setattr("agent_xlsx.adapters.xlwings_adapter.is_excel_available", lambda: True)

    result = resolve_engine("screenshot", "auto", libreoffice=True)
    assert result == "excel"


# ---------------------------------------------------------------------------
# Additional engine tests: all unavailable, "lo" alias, unknown engine
# ---------------------------------------------------------------------------


def test_all_engines_unavailable_raises(monkeypatch):
    """When all engines are unavailable, auto raises NoRenderingBackendError."""
    from agent_xlsx.utils.errors import NoRenderingBackendError

    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter._ASPOSE_AVAILABLE", None)
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter.is_aspose_available", lambda: False)
    monkeypatch.setattr("agent_xlsx.adapters.xlwings_adapter.is_excel_available", lambda: False)
    monkeypatch.setattr(
        "agent_xlsx.adapters.libreoffice_adapter.is_libreoffice_available",
        lambda: False,
    )

    with pytest.raises(NoRenderingBackendError):
        resolve_engine("screenshot", "auto", libreoffice=True)


def test_lo_alias_resolves_to_libreoffice(monkeypatch):
    """'lo' is accepted as alias for 'libreoffice'."""
    monkeypatch.setattr(
        "agent_xlsx.adapters.libreoffice_adapter.is_libreoffice_available",
        lambda: True,
    )

    result = resolve_engine("screenshot", "lo", libreoffice=True)
    assert result == "libreoffice"


def test_unknown_engine_raises(monkeypatch):
    """An unrecognised engine string raises an error."""
    from agent_xlsx.utils.errors import NoRenderingBackendError

    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter._ASPOSE_AVAILABLE", None)
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter.is_aspose_available", lambda: False)
    monkeypatch.setattr("agent_xlsx.adapters.xlwings_adapter.is_excel_available", lambda: False)

    with pytest.raises(NoRenderingBackendError):
        resolve_engine("screenshot", "nonexistent_engine", libreoffice=True)
