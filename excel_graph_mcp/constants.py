from pathlib import Path
from typing import Final

APP_NAME: Final = "ohh-my-excel"
VERSION: Final = "0.1.0"
GRAPH_DIR: Final = ".excel-graph"
DB_FILENAME: Final = "graph.db"
BACKUP_DIR: Final = "backups"

SCHEMA_VERSION: Final = 1

NODE_TYPES: Final[dict[str, str]] = {
    "WORKBOOK": "Workbook",
    "SHEET": "Sheet",
    "CELL": "Cell",
    "FORMULA": "Formula",
    "RANGE": "Range",
    "TABLE": "Table",
    "NAMED_RANGE": "NamedRange",
    "CHART": "Chart",
}

EDGE_TYPES: Final[dict[str, str]] = {
    "CONTAINS": "CONTAINS",
    "REFERENCES": "REFERENCES",
    "RANGE_REF": "RANGE_REF",
    "DEPENDS_ON": "DEPENDS_ON",
    "CROSS_SHEET_REF": "CROSS_SHEET_REF",
    "TABLE_REF": "TABLE_REF",
    "CHART_SOURCE": "CHART_SOURCE",
    "NAMED_REF": "NAMED_REF",
}

CONFIDENCE_LEVELS: Final[list[str]] = ["EXTRACTED", "INFERRED", "AMBIGUOUS"]

DEFAULT_MAX_FILE_SIZE_MB: Final = 100
DEFAULT_MAX_FILE_SIZE_BYTES: Final = DEFAULT_MAX_FILE_SIZE_MB * 1024 * 1024
DEFAULT_GENERATE_MAX_SIZE_MB: Final = 50
DEFAULT_BACKUP_RETENTION_DAYS: Final = 7
DEFAULT_MAX_BACKUPS: Final = 10
DEFAULT_WATCH_DEBOUNCE_MS: Final = 5000
DEFAULT_RATE_LIMIT: Final = 10

PERFORMANCE_TARGETS: Final[dict[str, float]] = {
    "build_initial_100_sheets_sec": 30.0,
    "incremental_update_ms": 500.0,
    "query_response_ms": 100.0,
    "search_response_ms": 10.0,
    "memory_mb": 200.0,
    "db_size_mb": 50.0,
}

def get_graph_dir(workbook_path: Path) -> Path:
    return workbook_path.parent / GRAPH_DIR

def get_db_path(workbook_path: Path) -> Path:
    return get_graph_dir(workbook_path) / DB_FILENAME

def get_backup_dir(workbook_path: Path) -> Path:
    return get_graph_dir(workbook_path) / BACKUP_DIR
