"""
Voice Bundle Database Access Facade.
====================================
Provides clean, decoupled database access points specifically for Voice Bundle
performance reporting workflows.

This module delegates directly to `services.shared.database`, exposing:
- `load_voice_metrics`: Fetches and normalizes voice purchases, call minute usages,
  and special 5 Birr package metrics across specified date ranges from Oracle or DuckDB.
"""

from services.shared.database import load_voice_data as load_voice_metrics

__all__ = ["load_voice_metrics"]
