"""Access to the bundled agent guides (the `SKILL.md` files).

Returns the guide text surfaced by the `guide` read op. Named `_guide` rather
than `_skill` because the package already ships a `_skill/` data directory
holding the Markdown files — a module of the same name can't coexist with it.

Variants that ship:

    _skill/wordcom-mcp/SKILL.md      (MCP tool form: read / exec / snapshot / write)
    _skill/wordcom-python/SKILL.md   (Python-API form)
"""

from __future__ import annotations

from importlib.resources import files

# kind -> installed skill directory name (also the `name:` in each frontmatter).
SKILLS: dict[str, str] = {
    "mcp": "wordcom-mcp",
    "python": "wordcom-python",
}


def skill_name(kind: str = "mcp") -> str:
    """The skill's canonical name / install directory (e.g. ``wordcom-mcp``)."""
    try:
        return SKILLS[kind]
    except KeyError as e:
        raise ValueError(f"unknown skill kind {kind!r}; expected one of {sorted(SKILLS)}") from e


def bundled_skill(kind: str = "mcp") -> str:
    """The packaged agent skill (SKILL.md) text, frontmatter and all."""
    name = skill_name(kind)
    return (files("wordcom") / "_skill" / name / "SKILL.md").read_text(encoding="utf-8")


def strip_frontmatter(md: str) -> str:
    """Drop a leading YAML frontmatter block (--- … ---), if present.

    Each bundled SKILL.md opens with `name:` / `description:` frontmatter for the
    agent-skill loader. That metadata is noise when the doc is served as a tool result
    or resource, so callers emit just the Markdown body.
    """
    lines = md.splitlines()
    if lines and lines[0].strip() == "---":
        for i in range(1, len(lines)):
            if lines[i].strip() == "---":
                return "\n".join(lines[i + 1 :]).lstrip("\n")
    return md


def skill_body(kind: str = "mcp") -> str:
    """The bundled guide with its YAML frontmatter stripped."""
    return strip_frontmatter(bundled_skill(kind))
