"""Server configuration: which apps to expose and the limits every tool obeys."""

from __future__ import annotations

from dataclasses import dataclass

# Every app the server knows about, in the order tools are listed.
APP_NAMES: tuple[str, ...] = ("word", "excel", "ppt")

# Apps whose tools are implemented today. The others are accepted names that
# fail with a clear "not available yet" message.
IMPLEMENTED_APPS: tuple[str, ...] = ("word", "excel", "ppt")

# Apps enabled when `--apps` isn't given. An app joins this list once it has
# passed a live agent test; until then it is opt-in (`--apps word,excel`).
DEFAULT_APPS: tuple[str, ...] = ("word",)

# Claude Code cuts tool output at about 25,000 tokens (MAX_MCP_OUTPUT_TOKENS)
# and warns above 10,000. 60,000 characters is roughly 15,000 tokens of JSON,
# which leaves room for the model's own reply.
DEFAULT_MAX_OUTPUT_CHARS = 60_000

# A rendered page's long edge, in pixels. About 1,200 keeps body text legible
# while holding each image to a predictable token cost.
DEFAULT_SNAPSHOT_MAX_DIM = 1200

# Pages returned by one snapshot call. More pages are listed in a note telling
# the model which `pages` range to ask for next.
DEFAULT_SNAPSHOT_MAX_PAGES = 10


@dataclass(frozen=True)
class ServerConfig:
    """Settings shared by every app's tools.

    Attributes:
        apps: App names to expose (see `APP_NAMES`).
        read_only: Register only tools that never change a document or a file.
        save_dirs: Directories saves and exports may write into. Empty means
            saving is off (default-deny).
        image_dirs: Optional allowlist for image files inserted from disk.
            Non-local paths (UNC shares, URLs) are always refused.
        max_output_chars: Longest text a tool returns before it is truncated
            with a note on how to narrow the request.
        snapshot_max_dim: Default cap on a rendered page's long edge in pixels.
        snapshot_max_pages: Most pages one snapshot call returns.
    """

    apps: tuple[str, ...] = DEFAULT_APPS
    read_only: bool = False
    save_dirs: tuple[str, ...] = ()
    image_dirs: tuple[str, ...] = ()
    max_output_chars: int = DEFAULT_MAX_OUTPUT_CHARS
    snapshot_max_dim: int = DEFAULT_SNAPSHOT_MAX_DIM
    snapshot_max_pages: int = DEFAULT_SNAPSHOT_MAX_PAGES

    def __post_init__(self) -> None:
        unknown = [a for a in self.apps if a not in APP_NAMES]
        if unknown:
            raise ValueError(
                f"unknown app(s): {', '.join(unknown)}; choose from {', '.join(APP_NAMES)}"
            )
        missing = [a for a in self.apps if a not in IMPLEMENTED_APPS]
        if missing:
            raise ValueError(
                f"app(s) not available yet: {', '.join(missing)}; "
                f"available now: {', '.join(IMPLEMENTED_APPS)}"
            )
        if not self.apps:
            raise ValueError("at least one app must be enabled")
        if self.max_output_chars < 1000:
            raise ValueError("max_output_chars must be at least 1000")
        if self.snapshot_max_dim < 100:
            raise ValueError("snapshot_max_dim must be at least 100")
        if self.snapshot_max_pages < 1:
            raise ValueError("snapshot_max_pages must be at least 1")
