"""Getting text out of the process on any machine. Specification_v1.md §12.

The tool's own vocabulary is not ASCII: sheet 00's blocks are anchored with ``⟦…⟧``, the
notes use ``→`` and ``·``, and error messages quote all three back at the reader. Neither
character exists in **cp1252**, which is still what Python uses for ``stdout`` on Windows
whenever the output is redirected to a file or a pipe.

The failure that causes is a nasty one. Everything works — the workbook is read, the
blocks are written, the logs are complete — and then the *last* step, printing the
summary, raises ``UnicodeEncodeError``. The user sees a traceback and reasonably concludes
the run failed, when in fact the output is sitting on disk beside them.

So the streams are put into UTF-8 before anything is printed. Where that is impossible —
a console genuinely unable to accept it — the fall-back is ``backslashreplace`` rather
than ``replace``: ``\\u27e6`` is ugly, but it says an anchor was there. A ``?`` says
nothing was, and this tool does not turn information into nothing quietly.
"""

from __future__ import annotations

import sys

ENCODING = "utf-8"
FALLBACK_ERRORS = "backslashreplace"


def _reconfigure(stream) -> bool:
    """Put one stream into UTF-8. ``False`` where the stream cannot be reconfigured."""
    reconfigure = getattr(stream, "reconfigure", None)
    if reconfigure is None:                      # already-wrapped or captured stream
        return False
    try:
        reconfigure(encoding=ENCODING, errors=FALLBACK_ERRORS)
        return True
    except (ValueError, OSError, LookupError):
        try:
            reconfigure(errors=FALLBACK_ERRORS)  # keep the encoding, survive the write
        except (ValueError, OSError, LookupError):
            return False
        return False


def use_utf8() -> None:
    """Called once, before the first print. Safe to call on any platform."""
    for stream in (sys.stdout, sys.stderr):
        if stream is not None:
            _reconfigure(stream)


def safe(text: str, stream=None) -> str:
    """The same text, guaranteed to survive ``stream``'s encoding.

    A last resort for streams that could not be reconfigured at all — an embedded
    interpreter, or a captured pipe someone else owns.
    """
    encoding = getattr(stream or sys.stdout, "encoding", None) or ENCODING
    try:
        text.encode(encoding)
        return text
    except (UnicodeEncodeError, LookupError):
        return text.encode(encoding, FALLBACK_ERRORS).decode(encoding, "replace")
