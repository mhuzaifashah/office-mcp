"""The PowerPoint backend the MCP tools call: one COM thread, the path rules.

Every method resolves its presentation on the COM thread (by name, or the
active one). The service never quits PowerPoint or closes a presentation.
"""

from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from typing import Any

from ..comthread import ComThread, Runner, call_timeout_from_env
from ..paths import PathNotAllowed, SavePolicy, reject_nonlocal
from .dispatch import com_message, dispatch_read, run_batch
from .errors import PowerPointBusyError, PowerPointError, PptOpError
from .session import controller_for, find_presentation, get_app, presentations

__all__ = ["OPEN_EXTENSIONS", "SAVE_FORMATS", "PowerPointService"]

OPEN_EXTENSIONS = (".pptx", ".pptm", ".ppt", ".potx", ".ppsx")
SAVE_FORMATS = {".pptx": 24, ".pptm": 25}  # PpSaveAsFileType
_PDF = 32  # ppSaveAsPDF


def _translate(exc: Exception) -> Exception:
    if type(exc).__name__ == "com_error":
        return PowerPointError(com_message(exc))
    return exc


def _parse_slides(spec: str | None, count: int) -> list[int]:
    """`"4"` or `"2-5"` (1-based, inclusive) -> slide numbers; None -> every slide."""
    if not spec:
        return list(range(1, count + 1))
    text = str(spec).strip()
    try:
        if "-" in text:
            a, b = (int(x) for x in text.split("-", 1))
        else:
            a = b = int(text)
    except ValueError:
        raise PptOpError(f'slides must look like "4" or "2-5", got {spec!r}') from None
    if a < 1 or b < a or b > count:
        raise PptOpError(f"slides {spec!r} out of range: the deck has {count}")
    return list(range(a, b + 1))


class PowerPointService:
    """Read, edit, render and save the presentations open in PowerPoint.

    Args:
        save_dirs: Folders `save`, `save_as` and `export_pdf` may write into.
        image_dirs: If set, pictures may only be inserted from these folders.
        runner: Where COM calls run; tests pass an `InlineRunner`.
    """

    def __init__(
        self,
        *,
        save_dirs: tuple[str, ...] = (),
        image_dirs: tuple[str, ...] = (),
        runner: Runner | None = None,
    ) -> None:
        self._policy = SavePolicy(save_dirs)
        self._image_dirs = tuple(Path(d).expanduser().resolve() for d in image_dirs)
        self._runner = runner or ComThread(
            "powerpoint",
            busy_error=PowerPointBusyError,
            closed_error=PowerPointError,
            timeout=call_timeout_from_env("OFFICE_MCP_CALL_TIMEOUT"),
        )

    def _run(self, fn: Any) -> Any:
        def job() -> Any:
            try:
                return fn()
            except Exception as exc:  # noqa: BLE001 - translated, then re-raised
                translated = _translate(exc)
                if translated is exc:
                    raise
                raise translated from exc

        return self._runner.run(job)

    def close(self) -> None:
        shutdown = getattr(self._runner, "shutdown", None)
        if shutdown is not None:
            shutdown()

    def screen_picture(self, raw: str) -> None:
        """A picture path must be local, and inside `image_dirs` when those are set."""
        reject_nonlocal(raw, what="picture")
        if self._image_dirs:
            resolved = Path(raw).expanduser().resolve()
            if not any(resolved == d or resolved.is_relative_to(d) for d in self._image_dirs):
                allowed = ", ".join(str(d) for d in self._image_dirs)
                raise PathNotAllowed(f"{raw!r} is outside the allowed image folders ({allowed})")

    # -- tools -------------------------------------------------------------------

    def read(self, command: str, *, presentation: str | None = None, **params: Any) -> Any:
        return self._run(lambda: dispatch_read(find_presentation(presentation), command, params))

    def exec(self, ops: list[dict[str, Any]], *, presentation: str | None = None) -> dict[str, Any]:
        return self._run(
            lambda: run_batch(find_presentation(presentation), ops, screen_path=self.screen_picture)
        )

    def snapshot(
        self,
        *,
        presentation: str | None = None,
        slides: str | None = None,
        max_dim: int = 1200,
        shape_ids: bool = True,
        max_slides: int = 10,
    ) -> tuple[list[tuple[int, bytes]], int]:
        """`([(slide, png), ...], total)`: at most `max_slides` rendered of `total` asked."""

        def job() -> tuple[list[tuple[int, bytes]], int]:
            pres = find_presentation(presentation)
            wanted = _parse_slides(slides, int(pres.Slides.Count))
            ctrl = controller_for(pres)
            w = float(pres.PageSetup.SlideWidth) or 960.0
            h = float(pres.PageSetup.SlideHeight) or 540.0
            scale = max_dim / max(w, h)
            width, height = max(1, round(w * scale)), max(1, round(h * scale))
            tmp = tempfile.mkdtemp(prefix="office-mcp-ppt-")
            shots: list[tuple[int, bytes]] = []
            try:
                for n in wanted[:max_slides]:
                    out = Path(tmp) / f"slide-{n}.png"
                    ctrl.slide_snapshot(
                        n, file_path=str(out), with_bboxes=shape_ids, width=width, height=height
                    )
                    shots.append((n, out.read_bytes()))
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
            return shots, len(wanted)

        return self._run(job)

    def write(
        self,
        command: str,
        *,
        presentation: str | None = None,
        path: str | None = None,
        overwrite: bool = False,
    ) -> dict[str, Any]:
        if command == "open":
            return self._open(path)
        if command == "new":
            return self._run(self._new)
        if command == "save":
            return self._run(lambda: self._save(find_presentation(presentation)))
        if command == "save_as":
            target = self._target(path, command, tuple(SAVE_FORMATS), overwrite)
            return self._run(lambda: self._save_as(find_presentation(presentation), target))
        if command == "export_pdf":
            target = self._target(path, command, (".pdf",), overwrite)
            return self._run(lambda: self._export_pdf(find_presentation(presentation), target))
        raise PptOpError(f"unknown write command {command!r}")

    # -- write commands ----------------------------------------------------------

    def _target(
        self, path: str | None, command: str, extensions: tuple[str, ...], overwrite: bool
    ) -> Path:
        if not path:
            raise PptOpError(f"write command {command!r} needs path")
        target = self._policy.resolve(path)
        if target.suffix.lower() not in extensions:
            raise PptOpError(f"{command} writes {', '.join(extensions)} files, not {target.name}")
        if target.exists() and not overwrite:
            raise PptOpError(f"{target} already exists; pass overwrite=true to replace it")
        return target

    def _open(self, path: str | None) -> dict[str, Any]:
        if not path:
            raise PptOpError("write command 'open' needs path")
        reject_nonlocal(path, what="presentation")
        file = Path(path).expanduser().resolve()
        if file.suffix.lower() not in OPEN_EXTENSIONS:
            raise PptOpError(
                f"PowerPoint opens {', '.join(OPEN_EXTENSIONS)} files, not {file.name}"
            )
        if not file.is_file():
            raise FileNotFoundError(f"no such file: {file}")

        def job() -> dict[str, Any]:
            app = get_app(start=True)
            for p in presentations(app):
                if str(p.FullName).lower() == str(file).lower():
                    try:
                        p.Windows(1).Activate()
                    except Exception:  # noqa: BLE001
                        pass
                    return _info(p, already_open=True)
            pres = app.Presentations.Open(str(file), 0, 0, -1)  # ReadOnly, Untitled, WithWindow
            return _info(pres, already_open=False)

        return self._run(job)

    def _new(self) -> dict[str, Any]:
        app = get_app(start=True)
        pres = app.Presentations.Add(-1)  # WithWindow
        layout = pres.SlideMaster.CustomLayouts(1)  # the title-slide layout
        pres.Slides.AddSlide(1, layout)
        return _info(pres)

    def _save(self, pres: Any) -> dict[str, Any]:
        if not str(pres.Path):
            raise PptOpError(
                f"{pres.Name} has never been saved; use save_as with a path in an allowed folder"
            )
        self._policy.resolve(str(pres.FullName))
        pres.Save()
        return {"ok": True, "command": "save", "path": str(pres.FullName)}

    def _save_as(self, pres: Any, target: Path) -> dict[str, Any]:
        pres.SaveAs(str(target), SAVE_FORMATS[target.suffix.lower()])
        return {"ok": True, "command": "save_as", "path": str(target), "presentation": pres.Name}

    def _export_pdf(self, pres: Any, target: Path) -> dict[str, Any]:
        # SaveCopyAs leaves the presentation tied to its own file (SaveAs would
        # re-point it at the PDF); ExportAsFixedFormat fails under late binding.
        pres.SaveCopyAs(str(target), _PDF)
        return {"ok": True, "command": "export_pdf", "path": str(target)}


def _info(pres: Any, *, already_open: bool | None = None) -> dict[str, Any]:
    out: dict[str, Any] = {
        "ok": True,
        "presentation": str(pres.Name),
        "path": str(pres.FullName),
        "slides": int(pres.Slides.Count),
    }
    if already_open is not None:
        out["already_open"] = already_open
    return out
