"""PowerPoint live-control prototype via pywin32 COM."""

from .controller import PowerPointController
from .errors import PowerPointError

__all__ = ["PowerPointController", "PowerPointError"]
