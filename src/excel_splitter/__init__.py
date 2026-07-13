"""Excel splitter core package."""

from .engine import SplitEngine
from .models import OutputArtifact, SheetConfig, SplitJob, SplitResult, SplitSummary

__all__ = [
    "OutputArtifact",
    "SheetConfig",
    "SplitEngine",
    "SplitJob",
    "SplitResult",
    "SplitSummary",
]

__version__ = "0.1.0"
