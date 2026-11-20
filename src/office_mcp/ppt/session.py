"""Find the presentation a tool call is about, in the PowerPoint the user has open.

Each call resolves its presentation afresh (by name, or the active one) on the
PowerPoint COM thread.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .controller import PowerPointController
from .errors import PowerPointNotRunningError, PresentationNotFoundError

PROG_ID = "PowerPoint.Application"


def get_app(*, start: bool = False) -> Any:
    """The running PowerPoint; with `start`, launch one when none is running."""
    import pywintypes  # noqa: PLC0415
    import win32com.client  # noqa: PLC0415

    try:
        return win32com.client.GetActiveObject(PROG_ID)
    except pywintypes.com_error:
        if not start:
            raise PowerPointNotRunningError("PowerPoint is not running") from None
    app = win32com.client.Dispatch(PROG_ID)
    try:
        app.Visible = -1  # msoTrue: PowerPoint can't run hidden with a window
    except Exception:  # noqa: BLE001
        pass
    return app


def presentations(app: Any) -> list[Any]:
    pres = app.Presentations
    return [pres(i) for i in range(1, int(pres.Count) + 1)]


def find_presentation(name: str | None = None) -> Any:
    """The presentation called `name` (file name, full path or stem), or the active one."""
    app = get_app()
    open_ = presentations(app)
    if not open_:
        raise PresentationNotFoundError("PowerPoint has no presentation open")
    if name:
        wanted = name.strip().lower()
        for p in open_:
            if str(p.Name).lower() == wanted or str(p.FullName).lower() == wanted:
                return p
        for p in open_:
            if Path(str(p.Name)).stem.lower() == wanted:
                return p
        raise PresentationNotFoundError(f"no open presentation is called {name!r}")
    try:
        return app.ActivePresentation
    except Exception:  # noqa: BLE001 - no active window (opened without one)
        return open_[0]


def controller_for(pres: Any) -> PowerPointController:
    """A controller bound to `pres` without `__init__` (which changes app visibility)."""
    ctrl = PowerPointController.__new__(PowerPointController)
    ctrl._app = pres.Application
    ctrl._pres = pres
    return ctrl
