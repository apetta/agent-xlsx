"""Memory management utilities."""

from __future__ import annotations

import psutil

from agent_xlsx.utils.constants import MAX_MEMORY_MB


def check_memory(limit_mb: float = MAX_MEMORY_MB) -> None:
    """Raise if current process exceeds memory budget."""
    from agent_xlsx.utils.errors import MemoryExceededError

    process = psutil.Process()
    memory_mb = process.memory_info().rss / 1024 / 1024
    if memory_mb > limit_mb:
        raise MemoryExceededError(memory_mb, limit_mb)


def get_memory_mb() -> float:
    """Return current process memory in MB."""
    return psutil.Process().memory_info().rss / 1024 / 1024
