"""CellPPT Studio Pro — Local Editable Academic PPTX Generation & Live PowerPoint Playback Product."""
from .nature2026_deck import generate_companion_files, generate_nature2026_pptx
from .paper_to_ppt import build_pptx_from_pdf
from .playback_engine import prepare_hidden_copy, run_live_playback
from .theme import THEMES, get_theme

__version__ = "2.0.0"
__all__ = [
    "THEMES",
    "get_theme",
    "generate_nature2026_pptx",
    "generate_companion_files",
    "build_pptx_from_pdf",
    "prepare_hidden_copy",
    "run_live_playback",
]
