from __future__ import annotations
from dataclasses import dataclass
import json
from typing import Optional, Tuple
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

@dataclass
class GeocodeConfig:
    enabled: bool = False
    provider: str = "stub"
    api_key: str = ""
    endpoint: str = ""
    timeout_seconds: int = 10

def build_address(street: str, zip_code: str, city: str, country: str) -> str:
    parts = []
    if street: parts.append(street)
    line2 = " ".join([p for p in [zip_code, city] if p])
    if line2: parts.append(line2)
    if country: parts.append(country)
    return ", ".join(parts)

def geocode_address(address: str, cfg: GeocodeConfig) -> Tuple[Optional[float], Optional[float], Optional[str]]:
    """Returns (lat, long, warning_message)."""
    if not address.strip():
        return (None, None, "Geocoding skipped because the address is blank.")

    if not cfg.enabled:
        return (None, None, "Geocoding disabled.")

    provider = (cfg.provider or "stub").strip().lower()
    if provider == "nominatim":
        endpoint = cfg.endpoint.strip() or "https://nominatim.openstreetmap.org/search"
        params = urlencode(
            {
                "q": address,
                "format": "jsonv2",
                "limit": 1,
            }
        )
        request = Request(
            f"{endpoint}?{params}",
            headers={
                "User-Agent": "Val-Model-Transposer/1.0",
                "Accept": "application/json",
            },
        )
        try:
            with urlopen(request, timeout=cfg.timeout_seconds) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except HTTPError as exc:
            return (None, None, f"Geocoding failed with HTTP {exc.code}.")
        except URLError as exc:
            return (None, None, f"Geocoding connection failed: {exc.reason}.")
        except TimeoutError:
            return (None, None, "Geocoding timed out.")
        except Exception as exc:
            return (None, None, f"Geocoding failed: {exc}")

        if not payload:
            return (None, None, "Geocoding found no result for the extracted address.")

        try:
            first_match = payload[0]
            return (float(first_match["lat"]), float(first_match["lon"]), None)
        except (KeyError, TypeError, ValueError, IndexError):
            return (None, None, "Geocoding returned an invalid response.")

    return (None, None, "Geocoding not implemented for provider: " + (cfg.provider or "unknown"))
