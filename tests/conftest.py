"""Root test config.

Importing ``office_mcp.word`` runs its sys.path bootstrap so the engine is
importable as top-level ``wordcom`` (needed by importlib.resources and tests).
"""

from __future__ import annotations

import office_mcp.word  # noqa: F401  (side effect: wordcom on sys.path)
