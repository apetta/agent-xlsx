"""agent-xlsx: XLSX file CLI built with Agent Experience (AX) in mind."""

# Harden stdlib XML parsers against XXE, entity expansion bombs, and DTD
# retrieval *before* any library (openpyxl, oletools, …) is imported.
import defusedxml

defusedxml.defuse_stdlib()

__version__ = "0.7.0"
