"""
Data Bundle Database Access Facade.
===================================
Provides clean, decoupled database access points specifically for Data Bundle
performance reporting workflows.

This module delegates directly to `services.shared.database`, exposing:
- `get_latest_sample_date`: Queries the latest available date key in local storage.
- `load_data_metrics`: Fetches and normalizes data bundle purchases and usage frames.
"""

from services.shared.database import get_latest_sample_date, load_bundle_data as load_data_metrics

__all__ = ["get_latest_sample_date", "load_data_metrics"]
