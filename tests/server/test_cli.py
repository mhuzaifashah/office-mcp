"""Command-line flags and their environment-variable twins."""

from __future__ import annotations

import os

import pytest

from office_mcp.server import cli
from office_mcp.server.config import DEFAULT_APPS, ServerConfig


def _config(argv, env=None):
    return cli.config_from_args(cli.build_parser().parse_args(argv), env or {})


def test_defaults():
    config = _config([])
    assert config == ServerConfig()
    assert config.apps == DEFAULT_APPS
    assert config.read_only is False
    assert config.save_dirs == ()


def test_flags():
    config = _config(
        [
            "--apps",
            "Word",
            "--read-only",
            "--save-dir",
            "out",
            "--save-dir",
            "pdf",
            "--image-dir",
            "img",
            "--max-output-chars",
            "5000",
            "--snapshot-max-pages",
            "3",
        ]
    )
    assert config.apps == ("word",)
    assert config.read_only is True
    assert config.save_dirs == ("out", "pdf")
    assert config.image_dirs == ("img",)
    assert config.max_output_chars == 5000
    assert config.snapshot_max_pages == 3


def test_env_vars_fill_in_missing_flags():
    env = {
        "OFFICE_MCP_APPS": "word",
        "OFFICE_MCP_READ_ONLY": "true",
        "OFFICE_MCP_SAVE_DIRS": os.pathsep.join(["a", "b"]),
        "OFFICE_MCP_IMAGE_DIRS": "img",
    }
    config = _config([], env)
    assert config.read_only is True
    assert config.save_dirs == ("a", "b")
    assert config.image_dirs == ("img",)


def test_flag_wins_over_env():
    config = _config(["--save-dir", "flag"], {"OFFICE_MCP_SAVE_DIRS": "env"})
    assert config.save_dirs == ("flag",)


def test_unknown_app_is_rejected():
    with pytest.raises(ValueError, match="unknown app"):
        _config(["--apps", "outlook"])


def test_every_app_can_be_enabled():
    assert _config(["--apps", "word,excel,ppt"]).apps == ("word", "excel", "ppt")


def test_version_flag(capsys):
    with pytest.raises(SystemExit) as exc:
        cli.main(["--version"])
    assert exc.value.code == 0
    assert capsys.readouterr().out.startswith("office-mcp ")
