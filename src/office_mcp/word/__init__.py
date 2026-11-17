"""Live Microsoft Word control over COM.

`WordController` (controller.py) is the public entry point; it's the action layer
over the `wordcom/` COM engine.
"""

from __future__ import annotations

from .controller import WordController

__all__ = ["WordController"]
