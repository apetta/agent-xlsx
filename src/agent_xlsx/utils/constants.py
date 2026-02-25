"""Token-efficiency caps and application constants."""

# Output capping limits
MAX_LOCATIONS = 20
MAX_FORMULA_CELLS = 50
MAX_FORMULA_PATTERNS = 10
MAX_SEARCH_RESULTS = 25
MAX_SEARCH_LIMIT = 1000  # Safety ceiling for user-specified --limit
MAX_VBA_LINES = 500
MAX_SAMPLE_ROWS = 10
DEFAULT_SAMPLE_ROWS = 3
MAX_READ_ROWS = 10_000

# Probe output truncation — cap string lengths to prevent token explosions
# on sheets with free-text columns (long descriptions, notes, etc.)
STRING_VALUE_MAX_CHARS = 80  # Cap for top_values in string_summary
SAMPLE_VALUE_MAX_CHARS = 100  # Cap for string values in sample rows
# Skip top_values in string_summary for columns with avg length above this
FREETEXT_AVG_LENGTH_THRESHOLD = 100

# Formula detection sampling — max cells to check via openpyxl when
# detecting uncached formulas in the default (Polars) read path
FORMULA_CHECK_SAMPLE_SIZE = 100

# Search optimization: full-load-then-slice heuristic —
# for large files, loading the full sheet and slicing in memory is faster
# than fastexcel's skip_rows (which scans sequentially from row 0)
SEARCH_FULLLOAD_FILE_SIZE_THRESHOLD = 10 * 1024 * 1024  # 10MB
SEARCH_FULLLOAD_ROW_RATIO = 0.10  # Use full-load when requested rows < 10% of total

# Memory management
MAX_MEMORY_MB = 500
CHUNK_THRESHOLD_BYTES = 100 * 1024 * 1024  # 100MB — chunk reads above this
WRITE_SIZE_WARN_BYTES = 20 * 1024 * 1024  # 20MB — warn/block before loading for write ops
CHUNK_SIZE_ROWS = 100_000

# Supported file extensions
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls", ".ods"}
WRITABLE_EXTENSIONS = {".xlsx", ".xlsm"}
VBA_EXTENSIONS = {".xlsm", ".xlsb"}

# Default pagination
DEFAULT_LIMIT = 100
DEFAULT_OFFSET = 0

# Screenshot quality thresholds
MIN_CAPTURE_WIDTH = 100
MIN_CAPTURE_HEIGHT = 100
