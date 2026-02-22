"""Engine detection and selection for commands that support multiple backends."""

from __future__ import annotations


def resolve_engine(command: str, engine: str, *, libreoffice: bool = True) -> str:
    """Resolve the backend engine for *command*.

    Evaluates *engine* (``"auto"``, ``"excel"``, ``"aspose"``,
    ``"libreoffice"`` / ``"lo"``) and returns the concrete engine name.

    Auto-detection priority: Aspose → Excel → LibreOffice (when *libreoffice*
    is ``True``).

    Raises an :class:`~agent_xlsx.utils.errors.AgentExcelError` subclass when
    the requested or detected engine is unavailable.
    """
    from agent_xlsx.adapters.aspose_adapter import is_aspose_available
    from agent_xlsx.adapters.xlwings_adapter import is_excel_available
    from agent_xlsx.utils.errors import (
        AsposeNotInstalledError,
        ExcelRequiredError,
        LibreOfficeNotFoundError,
        NoRenderingBackendError,
    )

    engine_lower = engine.lower()

    if engine_lower == "excel":
        if not is_excel_available():
            raise ExcelRequiredError(command)
        return "excel"

    if engine_lower == "aspose":
        if not is_aspose_available():
            raise AsposeNotInstalledError()
        return "aspose"

    if engine_lower in ("libreoffice", "lo"):
        if not libreoffice:
            # This command has no LibreOffice adapter; treat it as an unsupported
            # engine name rather than a missing installation.
            raise ExcelRequiredError(command)
        from agent_xlsx.adapters.libreoffice_adapter import is_libreoffice_available

        if not is_libreoffice_available():
            raise LibreOfficeNotFoundError()
        return "libreoffice"

    if engine_lower == "auto":
        if is_aspose_available():
            return "aspose"
        if is_excel_available():
            return "excel"
        if libreoffice:
            from agent_xlsx.adapters.libreoffice_adapter import is_libreoffice_available

            if is_libreoffice_available():
                return "libreoffice"
            raise NoRenderingBackendError(command)
        raise ExcelRequiredError(command)

    # Unknown engine string — error reflects what the command actually supports.
    raise NoRenderingBackendError(command) if libreoffice else ExcelRequiredError(command)
