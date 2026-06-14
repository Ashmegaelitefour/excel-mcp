"""Path security: allowlist-based file access control.

Set EXCEL_MCP_ALLOWED_DIRS to a colon-separated list of directories the server
is permitted to read from and write to. If unset, all paths are allowed (backward
compatible) but a warning is printed at startup.

Example:
    export EXCEL_MCP_ALLOWED_DIRS="/home/alice/reports:/home/alice/exports"
"""

from __future__ import annotations

import os
import warnings

_ALLOWED_DIRS: list[str] | None = None


def _load_allowed_dirs() -> list[str] | None:
    raw = os.environ.get("EXCEL_MCP_ALLOWED_DIRS", "").strip()
    if not raw:
        return None
    return [os.path.realpath(p) for p in raw.split(":") if p.strip()]


def get_allowed_dirs() -> list[str] | None:
    global _ALLOWED_DIRS
    if _ALLOWED_DIRS is None:
        _ALLOWED_DIRS = _load_allowed_dirs()
    return _ALLOWED_DIRS


def validate_path(file_path: str, *, write: bool = False) -> str:
    """Resolve and validate a path against the allowlist.

    Returns the resolved absolute path if allowed.
    Raises PermissionError if the path is outside all allowed directories.
    """
    resolved = os.path.realpath(os.path.abspath(file_path))
    allowed = get_allowed_dirs()

    if allowed is None:
        # No allowlist configured — permissive mode, warn once at startup
        return resolved

    for allowed_dir in allowed:
        # Use os.path.commonpath to prevent traversal tricks
        try:
            if os.path.commonpath([resolved, allowed_dir]) == allowed_dir:
                return resolved
        except ValueError:
            # commonpath raises ValueError on Windows when paths are on different drives
            continue

    action = "write to" if write else "read"
    raise PermissionError(
        f"Access denied: cannot {action} '{resolved}'.\n"
        f"Allowed directories: {allowed}\n"
        f"Set EXCEL_MCP_ALLOWED_DIRS to include this path."
    )


def warn_if_unrestricted() -> None:
    """Print a startup warning when no allowlist is configured."""
    if get_allowed_dirs() is None:
        warnings.warn(
            "excel-mcp: EXCEL_MCP_ALLOWED_DIRS is not set. "
            "The server can access any file path. "
            "Set this env var to restrict access to specific directories.",
            stacklevel=2,
        )
