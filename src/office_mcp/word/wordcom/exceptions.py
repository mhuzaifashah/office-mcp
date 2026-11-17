"""Public exception taxonomy for wordcom."""

from __future__ import annotations

from typing import Any


class WordcomError(Exception):
    """Base class for all wordcom errors."""


class WordNotRunningError(WordcomError):
    """No running Word instance is available."""


class DocumentNotFoundError(WordcomError):
    """The requested document is not open in Word."""

    def __init__(self, name: str) -> None:
        super().__init__(f"document not found: {name!r}")
        self.name = name


class AnchorNotFoundError(WordcomError):
    """The requested anchor (bookmark / content control / heading) does not exist."""

    def __init__(self, kind: str, name: str, *, hint: str | None = None) -> None:
        message = f"{kind} not found: {name!r}"
        if hint:
            message += f"; {hint}"
        super().__init__(message)
        self.kind = kind
        self.name = name
        self.hint = hint


class StyleNotFoundError(AnchorNotFoundError):
    """The requested paragraph or character style is not defined in the document.

    Subclass of `AnchorNotFoundError` so it shares the same error code (2) and so
    `except AnchorNotFoundError` catches both bookmark-misses and style-misses.
    Retryable after re-reading `doc.styles.list()`.
    """

    def __init__(self, name: str) -> None:
        super().__init__("style", name)


class AmbiguousMatchError(WordcomError):
    """A find/replace pattern matched more than one occurrence without disambiguation.

    Carries the list of matches so callers (notably LLM drivers) can pick an
    `occurrence` index and retry.
    """

    def __init__(self, find: str, matches: list[dict[str, Any]]) -> None:
        super().__init__(
            f"{len(matches)} matches for {find!r}; pass --all or --occurrence N to disambiguate"
        )
        self.find = find
        self.matches = matches


class ReplaceVerificationError(WordcomError):
    """A resolved replacement target didn't match the located text — refused to write.

    wordcom verifies each find/replace target against the located match before
    writing, and raises this instead of overwriting the wrong span. It means the
    document shifted out from under the match between locating and writing —
    typically because an earlier edit in the same batch moved later text, or
    Track Changes left both the inserted and deleted runs in place (so offsets no
    longer line up). Re-read the text (`find` / `paragraphs`) and retry, or target
    the span directly by anchor id. Maps to the generic error code (1). Not
    retryable as-is — the same call drifts again.
    """

    def __init__(
        self, find: str, expected: str, resolved: str, *, anchor_id: str | None = None
    ) -> None:
        super().__init__(
            f"replacement target for {find!r} resolved to {resolved!r}, not the expected "
            f"{expected!r}; refusing to overwrite a span that shifted under the match "
            f"(re-read the text and retry, or target it by anchor id)"
        )
        self.find = find
        self.expected = expected
        self.resolved = resolved
        self.anchor_id = anchor_id


class ImageSourceError(WordcomError):
    """An image given to `insert_image` couldn't be turned into an embeddable file.

    Raised for a missing or unreadable path, malformed base64, or bytes whose
    format isn't a recognised raster image (PNG/JPEG/GIF/BMP/TIFF). It's a
    bad-input error — not a "named thing is missing" — so it maps to the
    generic error code (1) rather than reusing the anchor-not-found code.
    Not retryable: fix the input.
    """

    def __init__(self, message: str) -> None:
        super().__init__(message)


class EquationError(WordcomError):
    """An equation given to `insert_equation` couldn't be built into Word math.

    Raised for malformed input (no input dialect given, or more than one of
    `unicodemath`/`latex`/`mathml`), unparseable MathML, a LaTeX source the
    optional converter rejects, or a missing LaTeX→MathML backend
    (`pip install "wordcom[latex]"`). It's a bad-input / dependency problem —
    not a "named thing is missing" — so it maps to the generic error code (1),
    like its sibling `ImageSourceError`. Not retryable: fix the input (or install
    the extra).
    """

    def __init__(self, message: str) -> None:
        super().__init__(message)


class PathNotAllowedError(WordcomError):
    """A filesystem path was refused by the gated surface's policy.

    wordcom's Python API is trusted and ungated, but the gated surface —
    whose inputs can be prompt-injected — run every filesystem path through a
    default-deny policy. This is raised when a **save / save-as / export-pdf**
    target falls outside the configured save-directory whitelist (or no whitelist
    is configured, so saving is off), or when an **image-source path** is a
    non-local form (UNC `\\…`, `file://`, or a URL) or sits outside the optional
    image-directory allowlist. It's a policy denial / bad-input error — not a
    "named thing is missing" — so it maps to the generic error code (1), keeping
    the six-code contract untouched. Not retryable: configure a whitelist
    (`--save-dir` / `WORDCOM_SAVE_DIRS`, `--image-dir` / `WORDCOM_IMAGE_DIRS`)
    or pass a local path inside it.
    """

    def __init__(self, message: str) -> None:
        super().__init__(message)


class SnapshotError(WordcomError):
    """A page/section snapshot couldn't be rendered.

    Raised when the optional PDF-rendering backend (PyMuPDF) isn't installed, or
    when rasterising the exported PDF fails. The PDF export itself goes through
    Word's COM, so a busy/modal Word surfaces as `WordBusyError`, not this. It's
    an environment/dependency problem rather than a "named thing is missing", so
    it maps to the generic error code (1). Fix by installing the extra:
    `pip install "wordcom[snapshot]"` (or `uv add "wordcom[snapshot]"`).
    """

    def __init__(self, message: str) -> None:
        super().__init__(message)


class ExcelNotAvailableError(WordcomError):
    """`insert_chart` needs Microsoft Excel, but it isn't available over COM.

    Charts are Excel-backed: `Range.InlineShapes.AddChart2` embeds a chart whose
    data lives in a hidden Excel workbook, so a working `Excel.Application` COM
    server is required. Raised (after a non-mutating probe, before any chart is
    inserted — the document is left untouched) when Excel isn't installed or
    can't be launched. Unlike `WordBusyError` this is a one-time environment
    problem, not a transient busy state — install Excel, or render the chart
    elsewhere and embed it with `insert_image`. It has its own error code (6),
    parallel to `WordNotRunningError`'s exit 4, so a missing-Excel failure is
    distinguishable from generic bad input. Not retryable.
    """

    def __init__(
        self, message: str = "Microsoft Excel is not available; charts require Excel installed"
    ) -> None:
        super().__init__(message)


class WordBusyError(WordcomError):
    """Word rejected the RPC — typically a modal dialog or a transient busy state.

    Retryable in principle; caller decides.
    """

    def __init__(
        self, message: str = "Word is busy or in a modal dialog", *, hresult: int | None = None
    ) -> None:
        super().__init__(message)
        self.hresult = hresult
        self.retryable = True


class OpError(WordcomError):
    """A batch/exec op (or a single dispatched write) was malformed.

    Raised for an unknown op kind, a missing required field, or a mutually
    exclusive pair given together (e.g. both `path` and `base64` for an image).
    It's a bad-input error — fix the request — so it maps to the generic exit
    code (1). Not retryable.
    """

    def __init__(self, message: str) -> None:
        super().__init__(message)


class ComError(WordcomError):
    """Generic wrapper for an unclassified pywintypes.com_error."""

    def __init__(
        self, message: str, *, hresult: int | None = None, description: str | None = None
    ) -> None:
        super().__init__(message)
        self.hresult = hresult
        self.description = description


# HRESULTs we recognise as "Word is momentarily unavailable" rather than a real error.
_BUSY_HRESULTS: frozenset[int] = frozenset(
    {
        0x80010001,  # RPC_E_CALL_REJECTED — call rejected by callee (modal dialog, busy)
        0x8001010A,  # RPC_E_SERVERCALL_RETRYLATER — server busy, retry later
        0x80010005,  # RPC_E_SERVERCALL_REJECTED — server rejected the call
        -2147418111,  # signed form of RPC_E_CALL_REJECTED
        -2147417846,  # signed form of RPC_E_SERVERCALL_RETRYLATER
    }
)


def _decode_com_error(exc: Any) -> tuple[int | None, str | None, str]:
    """Pull (hresult, description, readable_message) out of a pywintypes.com_error.

    pywintypes.com_error.args is (hresult, message, exc_info, arg_err) where exc_info,
    when present, is (wcode, source, description, helpfile, helpcontext, scode).
    """
    args: tuple[Any, ...] = getattr(exc, "args", ()) or ()
    hresult: int | None = None
    description: str | None = None
    message = str(exc)

    if len(args) >= 1 and isinstance(args[0], int):
        hresult = args[0]
    if len(args) >= 3 and args[2]:
        exc_info = args[2]
        try:
            description = exc_info[2] if len(exc_info) > 2 else None
            scode = exc_info[5] if len(exc_info) > 5 else None
        except (TypeError, IndexError):
            description, scode = None, None
        if scode is not None and hresult is None:
            hresult = scode

    parts = []
    if description:
        parts.append(description.strip())
    if hresult is not None:
        parts.append(f"HRESULT 0x{hresult & 0xFFFFFFFF:08X}")
    if parts:
        message = " — ".join(parts)
    return hresult, description, message


def from_com_error(exc: Any) -> WordcomError:
    """Classify a pywintypes.com_error into the appropriate wordcom exception."""
    hresult, description, message = _decode_com_error(exc)
    if hresult is not None and hresult in _BUSY_HRESULTS:
        return WordBusyError(message, hresult=hresult)
    return ComError(message, hresult=hresult, description=description)


def classify(exc: WordcomError) -> tuple[str, bool]:
    """Map a WordcomError to a stable `(code, retryable)` pair.

    The single source of truth for how a failure is labelled to callers: the
    dispatch layer surfaces `code` and `retryable` in its error reporting. Codes
    are `anchor_not_found` / `style_not_found`, `ambiguous_match`, `word_busy`,
    `word_not_running`, and `error` (everything else, including
    replace_verification).

    `isinstance` is checked subclass-first, so `StyleNotFoundError` gets its own
    `style_not_found` code (checked before its `AnchorNotFoundError` base).
    Unknown subclasses fall through to the generic `("error", False)`.
    """
    if isinstance(exc, StyleNotFoundError):
        return "style_not_found", False
    if isinstance(exc, AnchorNotFoundError):
        return "anchor_not_found", False
    if isinstance(exc, AmbiguousMatchError):
        return "ambiguous_match", True
    if isinstance(exc, ReplaceVerificationError):
        return "replace_verification", False
    if isinstance(exc, WordBusyError):
        return "word_busy", True
    if isinstance(exc, WordNotRunningError):
        return "word_not_running", False
    if isinstance(exc, ExcelNotAvailableError):
        return "excel_not_available", False
    if isinstance(exc, DocumentNotFoundError):
        return "document_not_found", False
    if isinstance(exc, PathNotAllowedError):
        return "path_not_allowed", False
    return "error", False
