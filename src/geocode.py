from __future__ import annotations
from dataclasses import dataclass
from typing import Optional, Tuple

@dataclass
class GeocodeConfig:
    enabled: bool = False
    provider: str = "stub"
    api_key: str = ""
    endpoint: str = ""

def build_address(street: str, zip_code: str, city: str, country: str) -> str:
    parts = []
    if street: parts.append(street)
    line2 = " ".join([p for p in [zip_code, city] if p])
    if line2: parts.append(line2)
    if country: parts.append(country)
    return ", ".join(parts)

def geocode_address(address: str, cfg: GeocodeConfig) -> Tuple[Optional[float], Optional[float], Optional[str]]:
    """
    Stub geocoder.
    Returns (lat, long, warning_message).
    Plug in a provider (e.g., Nominatim, Google, Mapbox) later.
    """
    if not cfg.enabled:
        return (None, None, "Geocoding disabled (stub).")
    # Placeholder for future implementation
    return (None, None, "Geocoding not implemented for provider: " + (cfg.provider or "unknown"))
