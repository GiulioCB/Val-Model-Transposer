from __future__ import annotations
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple, Any, Literal

YesNo = Literal["Yes", "No"]

@dataclass
class FilterQuestion:
    idx: int  # 1..16
    title: str
    required: bool
    kind: str  # 'dropdown' | 'text' | 'yesno'
    options: List[str] | None = None

@dataclass
class FilterAnswers:
    # keys: 1..16
    values: Dict[int, Any]

@dataclass
class StaticFields:
    property_name: Any
    street_address: Any
    country_region: Any
    city_town: Any
    zip_post_code: Any
    property_type: Any
    location: Any
    food_bev_operator: Any
    operator: Any
    rooms: Any
    chain_chainid: Any
    classification: Any
    management_company: Any
    owner_company: Any
    year_opened: Any
    meeting_space_sqm: Any
    meeting_rooms: Any
    meeting_max_capacity: Any
    ski: Any
    spa: Any
    healthclub: Any
    golf: Any
    boutique: Any
    food_outlets: Any
    beverage_outlets: Any
    pcd2: Any
    lat: Any
    long: Any
    currency: Any
