"""Database entry point used by voice performance reporting."""

from services.shared.database import load_voice_data as load_voice_metrics

__all__ = ["load_voice_metrics"]