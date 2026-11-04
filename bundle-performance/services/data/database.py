"""Database entry points used by data performance reporting."""

from services.shared.database import get_latest_sample_date, load_bundle_data

__all__ = ["get_latest_sample_date", "load_bundle_data"]
