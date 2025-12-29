from typing import Any, Dict
from .sqlserver_repository import SqlServerRepository
from .postgres_repository import PostgresRepository

def create_repo(conn_str: str, params: Dict[str, Any]) -> Any:
    # Prefer explicit dbKind param; fallback to sniffing connection string
    db_kind = (params.get("dbKind") or "").lower()
    if db_kind == "postgres" or conn_str.lower().startswith("postgres://") or "host=" in conn_str:
        return PostgresRepository(conn_str)
    return SqlServerRepository(conn_str)