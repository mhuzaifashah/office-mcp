"""Path rules for the Excel and PowerPoint tools.

Paths reach the tools from a model, so they are treated as untrusted:

- Network paths (UNC `\\\\host\\share`, `//host`) and URLs (`http:`, `file:`)
  are refused by a pure string check, before any filesystem call: just
  resolving a UNC path makes Windows authenticate to the remote host.
- Files are written only inside the folders the server was started with
  (`--save-dir`). With none, writing is off.

The same rules as Word's `wordcom.runtime._path_policy`, without importing
the Word engine.
"""

from __future__ import annotations

import os
import re
from collections.abc import Iterable
from pathlib import Path

__all__ = ["PathNotAllowed", "SavePolicy", "reject_nonlocal"]

# A scheme of 2+ characters and a colon; a drive letter (`C:`) has one, so it passes.
_SCHEME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9+.\-]+:")


class PathNotAllowed(Exception):
    """A path the tools refuse to touch."""


def reject_nonlocal(raw: str | os.PathLike[str], *, what: str = "file") -> None:
    """Raise if `raw` is a network path or a URL. No filesystem access."""
    s = str(raw)
    if s.startswith("\\\\") or s.startswith("//"):
        raise PathNotAllowed(f"refusing a network {what} path ({s!r}); use a local path")
    if _SCHEME_RE.match(s):
        raise PathNotAllowed(f"refusing a URL as a {what} path ({s!r}); use a local path")


class SavePolicy:
    """Default-deny list of folders files may be written into."""

    def __init__(self, save_dirs: Iterable[str | os.PathLike[str]] = ()) -> None:
        self._dirs = tuple(Path(d).expanduser().resolve() for d in save_dirs)

    @property
    def enabled(self) -> bool:
        return bool(self._dirs)

    def resolve(self, target: str | os.PathLike[str]) -> Path:
        """`target` as an absolute path inside an allowed folder, or raise.

        Resolves first (collapsing `..`, following links) so a crafted path
        can't escape, then checks containment.
        """
        reject_nonlocal(target, what="save target")
        if not self._dirs:
            raise PathNotAllowed(
                "saving is off: the server was started without --save-dir "
                f"(refused {str(target)!r})"
            )
        resolved = Path(target).expanduser().resolve()
        for d in self._dirs:
            if resolved == d or resolved.is_relative_to(d):
                return resolved
        allowed = ", ".join(str(d) for d in self._dirs)
        raise PathNotAllowed(f"{str(resolved)!r} is outside the allowed folders ({allowed})")
