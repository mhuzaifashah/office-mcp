"""Shared exception type. Lives in its own module so mixins can import it
without creating a circular dependency back to controller.py."""


class PowerPointError(Exception):
    """Raised when a PowerPoint operation fails."""
