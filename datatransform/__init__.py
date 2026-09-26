"""Data_Transform — metadata-driven extraction from Excel treaty submissions.

See specification_v1.md for the rules this package implements.
"""

from .model import Block, Confidence, Dataset, ExtractionError, Orientation, Record
from .runner import RunReport, run

__all__ = [
    "Block", "Confidence", "Dataset", "ExtractionError", "Orientation",
    "Record", "RunReport", "run",
]
from .constants import VERSION

__version__ = VERSION
