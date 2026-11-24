"""`msoffice-mcp` command line: parse flags, then serve over stdio.

Every flag has an environment-variable twin so MCP client configs can use
either. A flag wins over its variable. Directory lists in variables are
separated by `;` on Windows (`os.pathsep`).
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from collections.abc import Mapping, Sequence

from .. import __version__
from .config import (
    APP_NAMES,
    DEFAULT_APPS,
    DEFAULT_MAX_OUTPUT_CHARS,
    DEFAULT_SNAPSHOT_MAX_DIM,
    DEFAULT_SNAPSHOT_MAX_PAGES,
    ServerConfig,
)

__all__ = ["build_parser", "config_from_args", "main"]

ENV_APPS = "OFFICE_MCP_APPS"
ENV_READ_ONLY = "OFFICE_MCP_READ_ONLY"
ENV_SAVE_DIRS = "OFFICE_MCP_SAVE_DIRS"
ENV_IMAGE_DIRS = "OFFICE_MCP_IMAGE_DIRS"

_TRUE = {"1", "true", "yes", "on"}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="msoffice-mcp",
        description=(
            "MCP server for live Microsoft Office (Word today; Excel and PowerPoint next). "
            "Talks MCP over stdio; add it to your MCP client rather than running it by hand."
        ),
    )
    parser.add_argument(
        "--apps",
        metavar="LIST",
        help=(
            f"comma-separated apps to expose: {', '.join(APP_NAMES)} "
            f"(default: {','.join(DEFAULT_APPS)}; env {ENV_APPS})"
        ),
    )
    parser.add_argument(
        "--read-only",
        action="store_true",
        default=None,
        help=f"only register tools that never change a document or file (env {ENV_READ_ONLY}=1)",
    )
    parser.add_argument(
        "--save-dir",
        action="append",
        metavar="DIR",
        help=(
            "directory saves and exports may write into; repeat for more. "
            f"Without one, saving is off (env {ENV_SAVE_DIRS})"
        ),
    )
    parser.add_argument(
        "--image-dir",
        action="append",
        metavar="DIR",
        help=(
            "only insert images from disk inside this directory; repeat for more "
            f"(env {ENV_IMAGE_DIRS})"
        ),
    )
    parser.add_argument(
        "--max-output-chars",
        type=int,
        default=DEFAULT_MAX_OUTPUT_CHARS,
        metavar="N",
        help="longest text one tool call returns before it is truncated (default: %(default)s)",
    )
    parser.add_argument(
        "--snapshot-max-dim",
        type=int,
        default=DEFAULT_SNAPSHOT_MAX_DIM,
        metavar="PX",
        help="default cap on a rendered page's long edge in pixels (default: %(default)s)",
    )
    parser.add_argument(
        "--snapshot-max-pages",
        type=int,
        default=DEFAULT_SNAPSHOT_MAX_PAGES,
        metavar="N",
        help="most pages one snapshot call returns (default: %(default)s)",
    )
    parser.add_argument(
        "--log-level",
        default="WARNING",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="log level for messages on stderr (default: %(default)s)",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def _split_dirs(value: str | None) -> list[str]:
    return [d.strip() for d in (value or "").split(os.pathsep) if d.strip()]


def config_from_args(
    args: argparse.Namespace, env: Mapping[str, str] | None = None
) -> ServerConfig:
    """Merge parsed flags with environment variables into a `ServerConfig`."""
    env = os.environ if env is None else env

    apps_raw = args.apps if args.apps is not None else env.get(ENV_APPS)
    apps = (
        tuple(a.strip().lower() for a in apps_raw.split(",") if a.strip())
        if apps_raw
        else DEFAULT_APPS
    )

    read_only = args.read_only
    if read_only is None:
        read_only = env.get(ENV_READ_ONLY, "").strip().lower() in _TRUE

    save_dirs = args.save_dir if args.save_dir else _split_dirs(env.get(ENV_SAVE_DIRS))
    image_dirs = args.image_dir if args.image_dir else _split_dirs(env.get(ENV_IMAGE_DIRS))

    return ServerConfig(
        apps=apps,
        read_only=bool(read_only),
        save_dirs=tuple(save_dirs),
        image_dirs=tuple(image_dirs),
        max_output_chars=args.max_output_chars,
        snapshot_max_dim=args.snapshot_max_dim,
        snapshot_max_pages=args.snapshot_max_pages,
    )


def main(argv: Sequence[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    # stdout carries the MCP protocol; every log line goes to stderr.
    logging.basicConfig(
        level=args.log_level,
        stream=sys.stderr,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )

    if sys.platform != "win32":
        parser.exit(
            2,
            "msoffice-mcp: Microsoft Office automation needs Windows. If you are in WSL, "
            "configure your MCP client to start the Windows build of msoffice-mcp.\n",
        )

    try:
        config = config_from_args(args)
    except ValueError as exc:
        parser.error(str(exc))

    from .build import build_server  # noqa: PLC0415 — keep --help/--version fast

    build_server(config).run(transport="stdio")
