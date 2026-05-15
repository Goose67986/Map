"""
Yard Sale Finder v6.5

Main cleanup from v5.5:
- Removed geopy/Nominatim completely.
- Uses Google Geocoding API only for address lookup and Craigslist coordinate reverse lookup.
- Uses one SaleEvent object instead of keeping separate parallel lists.
- Adds Google geocode cache to reduce repeat API calls.
- Keeps photo OCR, GSALR, Craigslist, AuctionZip, PennLive, Folium map, sale selection export,
  embedded image links, historical description matching, and JSON export.

Before running:
1) Put your Google Maps Geocoding API key in GOOGLE_MAPS_API_KEY below,
   or set an environment variable named GOOGLE_MAPS_API_KEY.
2) Optional: put your OCR.Space key in OCR_SPACE_API_KEY, or leave "helloworld" for demo testing.
"""

from __future__ import annotations

import base64
import html
import json
import os
import random
import re
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import quote_plus, urljoin

import folium
import pandas as pd
import requests
from bs4 import BeautifulSoup
from folium import Element
from rapidfuzz import fuzz


# =============================================================================
# CONFIG
# =============================================================================

VERSION = "6.5"

DOWNLOAD_DIR = Path("/storage/emulated/0/Download")
PHOTO_FOLDER = DOWNLOAD_DIR
OUTPUT_DIR = Path(".")
MAP_FILE = OUTPUT_DIR / "map.html"
EXPORT_JSON_FILE = OUTPUT_DIR / "yard_sales_latest.json"
PHOTO_OCR_JSON = DOWNLOAD_DIR / "yard_sales_from_photos.json"
GEOCODE_CACHE_FILE = OUTPUT_DIR / "google_geocode_cache.json"

# Safer than hard-coding your real key into the script.
# You can either:
#   1) replace "PUT_YOUR_GOOGLE_API_KEY_HERE"
#   2) or set an environment variable named GOOGLE_MAPS_API_KEY
GOOGLE_MAPS_API_KEY = os.environ.get("GOOGLE_MAPS_API_KEY", "PUT_YOUR_GOOGLE_API_KEY_HERE")
GOOGLE_MAPS_API_KEY = "AIzaSyBjDVzoi-JgNUaBKyRgMhxSqipgBrxWkG8"


OCR_SPACE_API_KEY = os.environ.get("OCR_SPACE_API_KEY", "helloworld")
OCR_SPACE_URL = "https://api.ocr.space/parse/image"

IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".webp")

REQUEST_TIMEOUT = 20
GOOGLE_SLEEP_SECONDS = 0.15
SCRAPE_SLEEP_SECONDS = 0.75

MAP_CENTER = [40.2379, -76.9223]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/128.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

KEYWORDS = ["Community", "Neighborhood", "Garage", "Yard", "Rummage", "Church", "Estate", "Development"]
LINKS_TO_REMOVE = ["gsalr.com", "List", "estatesales", "salesestate", "companies", "real estate", "news"]

HISTORICAL_EXCEL = Path("yardsale_descriptions.xlsx")
HISTORICAL_MATCH_SCORE = 87

HIGH_PRIORITY_NAME_KEYWORDS = ["community", "neighborhood", "flea", "development"]
MEDIUM_PRIORITY_KEYWORDS = ["multi"]
LOCATION_KEYWORDS = ["school", "church", "center", "building", "ymca"]
AUCTION_KEYWORDS = ["auction", "bidding"]


# =============================================================================
# MODELS
# =============================================================================

@dataclass
class SaleEvent:
    title: str = "No Sale Data"
    description: str = "No Sale Data"
    raw_address: str = ""
    date_text: Any = "No Times Found"
    source_link: str = ""
    source: str = ""
    source_file: str = ""

    formatted_address: str = ""
    latitude: Optional[float] = None
    longitude: Optional[float] = None
    google_place_id: str = ""
    google_maps_link: str = ""
    historical_details: List[Dict[str, str]] = field(default_factory=list)

    def has_location(self) -> bool:
        return self.latitude is not None and self.longitude is not None

    def formatted_date_html(self) -> str:
        if isinstance(self.date_text, (list, tuple)):
            return "<br>".join(html.escape(str(x)) for x in self.date_text if str(x).strip())
        return html.escape(str(self.date_text))


# =============================================================================
# GENERAL HELPERS
# =============================================================================

def log(message: str) -> None:
    print(message, flush=True)


def safe_text(tag: Any, default: str = "") -> str:
    return tag.get_text(" ", strip=True) if tag else default


def clean_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def get_next_saturday() -> datetime:
    today = datetime.today()
    days_until_saturday = (5 - today.weekday()) % 7
    return today + timedelta(days=days_until_saturday)


def build_main_urls() -> List[str]:
    next_saturday = get_next_saturday()
    formatted_date = next_saturday.strftime("%Y-%m-%d")
    year = next_saturday.strftime("%Y")
    month = next_saturday.strftime("%m")
    day = next_saturday.strftime("%d")

    return [
        "https://gsalr.com/garage-sales-harrisburg-pa.html?day=5",
        f"https://harrisburg.craigslist.org/search/gms?lat=40.2736&lon=-76.8847&sale_date={formatted_date}&search_distance=10",
        # Uncomment if wanted:
        # f"https://www.auctionzip.com/cgi-bin/auctionlist.cgi?txtSearchZip=17055&txtSearchRadius=30&idxSearchCategory=0&gid=0&year={year}&month={month}&day={day}&txtSearchKeywords=&showlive=1",
        # "https://classifieds.pennlive.com/pennlive/category/garage-sale-estate-sale-auctions/garage-yard-estate-sales",
        # "https://garagesalefinder.com/yard-sales/mechanicsburg-pa/",
        # "https://garagesalefinder.com/yard-sales/camp-hill-pa/",
    ]


def request_soup(url: str, retries: int = 3, timeout: int = REQUEST_TIMEOUT) -> Optional[BeautifulSoup]:
    for attempt in range(1, retries + 1):
        try:
            response = requests.get(url, headers=HEADERS, timeout=timeout)
            if response.status_code != 200:
                wait = random.uniform(1.5, 4.5)
                log(f"[{response.status_code}] {url} - retry {attempt}/{retries} in {wait:.1f}s")
                time.sleep(wait)
                continue
            return BeautifulSoup(response.text, "html.parser")
        except requests.RequestException as exc:
            wait = random.uniform(1.5, 4.5)
            log(f"Request failed: {exc} - retry {attempt}/{retries} in {wait:.1f}s")
            time.sleep(wait)
    return None


def load_json(path: Path, fallback: Any) -> Any:
    if not path.exists():
        return fallback
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return fallback


def save_json(path: Path, data: Any) -> None:
    path.write_text(json.dumps(data, indent=4, ensure_ascii=False), encoding="utf-8")


def address_contains_near(address: str) -> bool:
    return bool(re.search(r"\bnear\b", str(address or ""), flags=re.IGNORECASE))


def normalize_near_address(address: str) -> str:
    # Google understands "near" sometimes, but "and" is better for street intersections.
    return re.sub(r"\bnear\b", "and", str(address or ""), flags=re.IGNORECASE)


def maps_search_url(query: str) -> str:
    return "https://www.google.com/maps/search/?api=1&query=" + quote_plus(query)


def image_to_base64(path: str) -> Tuple[Optional[str], Optional[str]]:
    try:
        file_path = Path(path)
        ext = file_path.suffix.lower()
        mime = {
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".png": "image/png",
            ".webp": "image/webp",
        }.get(ext, "image/jpeg")

        encoded = base64.b64encode(file_path.read_bytes()).decode()
        return encoded, mime
    except Exception as exc:
        log(f"Could not encode image {path}: {exc}")
        return None, None


# =============================================================================
# GOOGLE GEOCODING ONLY
# =============================================================================

class GoogleGeocoder:
    def __init__(self, api_key: str, cache_file: Path = GEOCODE_CACHE_FILE) -> None:
        self.api_key = api_key
        self.cache_file = cache_file
        self.cache: Dict[str, Dict[str, Any]] = load_json(cache_file, {})

        if not self.api_key or self.api_key == "PUT_YOUR_GOOGLE_API_KEY_HERE":
            log("WARNING: GOOGLE_MAPS_API_KEY is not set. Geocoding will fail until you add your key.")

    def _cache_key(self, prefix: str, value: str) -> str:
        return f"{prefix}:{clean_spaces(value).lower()}"

    def _save_cache(self) -> None:
        save_json(self.cache_file, self.cache)

    def geocode_address(self, address: str) -> Optional[Dict[str, Any]]:
        address = clean_spaces(address)
        if not address:
            return None

        lookup_address = normalize_near_address(address)
        key = self._cache_key("address", lookup_address)

        if key in self.cache:
            return self.cache[key]

        params = {
            "address": lookup_address,
            "key": self.api_key,
        }

        try:
            response = requests.get(
                "https://maps.googleapis.com/maps/api/geocode/json",
                params=params,
                timeout=REQUEST_TIMEOUT,
            )
            data = response.json()
        except Exception as exc:
            log(f"Google geocode request failed for {address}: {exc}")
            return None

        result = self._parse_google_result(data, original_query=lookup_address)
        if result:
            self.cache[key] = result
            self._save_cache()
            time.sleep(GOOGLE_SLEEP_SECONDS)
        else:
            status = data.get("status", "UNKNOWN")
            msg = data.get("error_message", "")
            log(f"Google geocode failed for '{address}': {status} {msg}")

        return result

    def reverse_geocode(self, latitude: float, longitude: float) -> Optional[Dict[str, Any]]:
        latlng = f"{latitude:.6f},{longitude:.6f}"
        key = self._cache_key("latlng", latlng)

        if key in self.cache:
            return self.cache[key]

        params = {
            "latlng": latlng,
            "key": self.api_key,
        }

        try:
            response = requests.get(
                "https://maps.googleapis.com/maps/api/geocode/json",
                params=params,
                timeout=REQUEST_TIMEOUT,
            )
            data = response.json()
        except Exception as exc:
            log(f"Google reverse geocode request failed for {latlng}: {exc}")
            return None

        result = self._parse_google_result(data, original_query=latlng)
        if result:
            self.cache[key] = result
            self._save_cache()
            time.sleep(GOOGLE_SLEEP_SECONDS)
        else:
            status = data.get("status", "UNKNOWN")
            msg = data.get("error_message", "")
            log(f"Google reverse geocode failed for '{latlng}': {status} {msg}")

        return result

    @staticmethod
    def _parse_google_result(data: Dict[str, Any], original_query: str) -> Optional[Dict[str, Any]]:
        if data.get("status") != "OK" or not data.get("results"):
            return None

        first = data["results"][0]
        loc = first.get("geometry", {}).get("location", {})
        lat = loc.get("lat")
        lng = loc.get("lng")

        if lat is None or lng is None:
            return None

        formatted = first.get("formatted_address") or original_query
        return {
            "formatted_address": formatted,
            "latitude": float(lat),
            "longitude": float(lng),
            "place_id": first.get("place_id", ""),
        }


def geocode_events(events: List[SaleEvent], geocoder: GoogleGeocoder) -> List[SaleEvent]:
    unfound = []

    for event in events:
        raw_query = event.raw_address or event.formatted_address
        force_intersection_lookup = address_contains_near(raw_query)

        # Craigslist can provide approximate coordinates from the listing itself.
        # For "A near B" addresses, those coordinates may point to a guessed nearby address,
        # so force a fresh Google lookup using "A and B" as an intersection query.
        if event.has_location() and not force_intersection_lookup:
            continue

        if force_intersection_lookup:
            query = normalize_near_address(raw_query)
            log(f"Intersection lookup: {raw_query} -> {query}")
        else:
            query = event.formatted_address or event.raw_address

        result = geocoder.geocode_address(query)

        if result:
            event.formatted_address = result["formatted_address"]
            event.latitude = result["latitude"]
            event.longitude = result["longitude"]
            event.google_place_id = result.get("place_id", "")
            event.google_maps_link = maps_search_url(query if force_intersection_lookup else event.formatted_address)
            log(f"Geocoded: {query} -> {event.formatted_address}")
        else:
            event.google_maps_link = maps_search_url(query)
            unfound.append(event)
            log(f"NOT FOUND: {event.title} - {query}")

    log(f"Geocoded {sum(1 for e in events if e.has_location())}/{len(events)} events.")
    if unfound:
        log(f"{len(unfound)} addresses were not found by Google.")

    return events


# =============================================================================
# HISTORICAL SALE MATCHING
# =============================================================================

def check_historical_sale_data(title: str, description: str, address: str) -> List[Dict[str, str]]:
    if not HISTORICAL_EXCEL.exists():
        return []

    full_text = f"{title} {description} {address}".lower()
    history: List[Dict[str, str]] = []

    try:
        xls = pd.ExcelFile(HISTORICAL_EXCEL)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            if df.shape[0] < 1 or df.shape[1] < 2:
                continue

            titles = df.iloc[:, 0].fillna("").astype(str)
            descriptions = df.iloc[:, 1].fillna("").astype(str)

            for old_title, old_desc in zip(titles, descriptions):
                old_title_clean = old_title.strip()
                if not old_title_clean:
                    continue

                if fuzz.partial_ratio(old_title_clean.lower(), full_text) >= HISTORICAL_MATCH_SCORE:
                    history.append({
                        "sheet": str(sheet_name),
                        "title": old_title_clean,
                        "description": old_desc.strip(),
                    })
    except Exception as exc:
        log(f"Historical lookup skipped: {exc}")

    return history


# =============================================================================
# LINK DISCOVERY
# =============================================================================

def get_relevant_links(url: str, keywords: Iterable[str], links_to_remove: Iterable[str]) -> List[Tuple[str, str]]:
    soup = request_soup(url)
    if not soup:
        return []

    relevant = []
    for link in soup.find_all("a"):
        href = link.get("href")
        text = clean_spaces(link.get_text(" "))
        if not href or not text:
            continue

        if any(k.lower() in text.lower() for k in keywords):
            if not any(bad.lower() in text.lower() for bad in links_to_remove):
                relevant.append((text, urljoin(url, href)))

    return list(dict.fromkeys(relevant))


def get_relevant_pennlive_links(url: str) -> List[Tuple[str, str]]:
    soup = request_soup(url)
    if not soup:
        return []

    links = []
    for link in soup.find_all("a"):
        href = link.get("href")
        text = clean_spaces(link.get_text(" "))
        if href and "show more" in text.lower():
            links.append((text, urljoin(url, href)))

    return list(dict.fromkeys(links))


def get_relevant_auctionzip_links(url: str) -> List[Tuple[str, str]]:
    soup = request_soup(url, retries=5)
    if not soup:
        return []

    relevant = []
    listings = soup.find_all("div", class_="az-ListOfLlisting-body")

    for listing in listings:
        a = listing.find("a", class_="az-ListOfLlisting-body__link")
        if not a or not a.get("href"):
            continue

        href = a["href"]
        if "(" in href and ")" in href:
            href = href[href.find("(") + 1:href.find(")")]

        if href.startswith("//"):
            href = "https:" + href
        elif href.startswith("/"):
            href = "https://www.auctionzip.com" + href

        title = clean_spaces(a.get_text(" ").replace("\xa0", " "))
        relevant.append((title or href, href))

    return list(dict.fromkeys(relevant))


# =============================================================================
# SCRAPERS
# =============================================================================

def scrape_gsalr(url: str) -> Optional[SaleEvent]:
    log(f"SCRAPING GSALR: {url}")
    soup = request_soup(url)
    if not soup:
        return None

    title = safe_text(soup.find(itemprop="name"), "No Sale Data")
    description = safe_text(soup.find(itemprop="description"), "No Sale Data")

    sale_times = []
    for time_div in soup.find_all("div", class_="sale-date-cards"):
        text = clean_spaces(time_div.get_text(" "))
        if text:
            text = re.sub(r"(PM|AM)", r"\1\n", text, flags=re.IGNORECASE).strip()
            sale_times.append(text)

    street = safe_text(soup.find(itemprop="streetAddress"), "")
    city = safe_text(soup.find(itemprop="addressLocality"), "")
    state = safe_text(soup.find(itemprop="addressRegion"), "")
    zipcode = safe_text(soup.find(itemprop="postalCode"), "")
    address = clean_spaces(f"{street}, {city}, {state} {zipcode}".strip(" ,"))

    return SaleEvent(
        title=title,
        description=description,
        raw_address=address,
        date_text=sale_times or "No Times Found",
        source_link=url,
        source="GSALR",
    )


def scrape_craigslist(url: str, geocoder: GoogleGeocoder) -> Optional[SaleEvent]:
    log(f"SCRAPING CRAIGSLIST: {url}")
    soup = request_soup(url)
    if not soup:
        return None

    title = ""
    description = ""
    city = ""
    state = ""
    formatted_from_coords = ""
    lat = None
    lng = None

    for tag in soup.find_all("meta"):
        attrs = {str(k).lower().strip(): v for k, v in tag.attrs.items()}
        content = attrs.get("content", "")

        if attrs.get("property") == "og:title":
            title = clean_spaces(content)
        elif attrs.get("property") == "og:description":
            description = clean_spaces(content)
        elif attrs.get("name") == "geo.region":
            state = str(content)[-2:]
        elif attrs.get("name") == "geo.placename":
            city = clean_spaces(content)
        elif attrs.get("name") == "geo.position":
            try:
                lat_str, lng_str = str(content).split(";")
                lat = round(float(lat_str), 6)
                lng = round(float(lng_str), 6)
            except Exception:
                pass

    if lat is not None and lng is not None:
        rev = geocoder.reverse_geocode(lat, lng)
        if rev:
            formatted_from_coords = rev["formatted_address"]

    mapaddress = safe_text(soup.find("div", class_="mapaddress"), "")
    if mapaddress:
        raw_address = clean_spaces(f"{mapaddress}, {city}, {state}".strip(" ,"))
    else:
        raw_address = formatted_from_coords

    sale_dates = [clean_spaces(d.get_text(" ")) for d in soup.find_all("a", class_="valu") if clean_spaces(d.get_text(" "))]

    start_time = ""
    sale_time_div = soup.find("div", class_="attr sale_time")
    if sale_time_div:
        start_time = safe_text(sale_time_div.find("span", class_="valu"), "")

    date_text = sale_dates
    if start_time:
        date_text = sale_dates + [start_time]

    event = SaleEvent(
        title=title or "Craigslist Sale",
        description=description,
        raw_address=raw_address,
        date_text=date_text or "No Dates Found",
        source_link=url,
        source="Craigslist",
    )

    # Craigslist may only expose approximate coords. Keep them if Google returned an address.
    if formatted_from_coords:
        event.formatted_address = formatted_from_coords
        event.latitude = lat
        event.longitude = lng
        event.google_maps_link = maps_search_url(formatted_from_coords)

    return event


def scrape_auctionzip(url: str) -> Optional[SaleEvent]:
    log(f"SCRAPING AUCTIONZIP: {url}")
    soup = request_soup(url, retries=5)
    if not soup:
        return None

    script_tag = soup.find("script", string=re.compile(r"dataLayer\.push"))
    if not script_tag or not script_tag.string:
        log("AuctionZip dataLayer JSON not found.")
        return None

    match = re.search(r"dataLayer\.push\((\{.*?\})\);", script_tag.string, re.DOTALL)
    if not match:
        log("AuctionZip dataLayer JSON pattern not matched.")
        return None

    try:
        data = json.loads(match.group(1))
    except Exception as exc:
        log(f"AuctionZip JSON parse failed: {exc}")
        return None

    title = clean_spaces(str(data.get("name") or "Auction"))
    sale_date = clean_spaces(str(data.get("SaleDate") or "No Dates Found"))
    location = clean_spaces(str(data.get("sellerLocation") or ""))
    desc_tag = soup.find("meta", attrs={"name": "description"})
    description = clean_spaces(desc_tag.get("content", "")) if desc_tag else ""

    return SaleEvent(
        title=title.replace(",", ""),
        description=description.replace("/", ""),
        raw_address=location.replace(",", ""),
        date_text=sale_date.replace(",", ""),
        source_link=url,
        source="AuctionZip",
    )


def scrape_pennlive(url: str) -> Optional[SaleEvent]:
    log(f"SCRAPING PENNLIVE: {url}")
    soup = request_soup(url)
    if not soup:
        return None

    main_ad = soup.find("div", class_="sr_ad_frame")
    if not main_ad:
        log("PennLive main ad block not found.")
        return None

    title = safe_text(main_ad.find("span", class_="sr_ad_title"), "PennLive Sale")
    description = safe_text(main_ad.find("p", itemprop="description"), "")

    date_match = re.search(
        r"\b(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s+([A-Za-z]+ \d{1,2})",
        description,
        re.IGNORECASE,
    )
    day = f"{date_match.group(1).capitalize()}, {date_match.group(2)}" if date_match else "N/A"

    time_match = re.search(r"\b\d{1,2}\s*(?:am|pm)?\s*(?:-|to)\s*\d{1,2}\s*(?:am|pm)?\b", description, re.IGNORECASE)
    time_text = time_match.group(0).replace("to", "-") if time_match else "N/A"

    street_match = re.search(
        r"\d{2,6}\s+[\w\s.'-]+(?:Street|St|Road|Rd|Avenue|Ave|Court|Ct|Drive|Dr|Lane|Ln|Boulevard|Blvd|Way|Pike)\.?",
        description,
        re.IGNORECASE,
    )
    street = street_match.group(0).strip() if street_match else ""

    city_match = re.search(r"[A-Za-z\s.'-]+,\s*PA\s*\d{5}|New Cumberland,?\s*\d{5}?", description, re.IGNORECASE)
    city = city_match.group(0).strip() if city_match else ""

    address = clean_spaces(f"{street}, {city}".strip(" ,"))

    return SaleEvent(
        title=title,
        description=description,
        raw_address=address,
        date_text=[day, time_text],
        source_link=url,
        source="PennLive",
    )


def scrape_site(url: str, geocoder: GoogleGeocoder) -> List[SaleEvent]:
    if "pennlive" in url.lower():
        links = get_relevant_pennlive_links(url)
    elif "auctionzip" in url.lower():
        links = get_relevant_auctionzip_links(url)
    else:
        links = get_relevant_links(url, KEYWORDS, LINKS_TO_REMOVE)

    if not links:
        log(f"No relevant links found at {url}.")
        return []

    log(f"Found {len(links)} relevant links at {url}.")
    events: List[SaleEvent] = []

    for _, link in links:
        event = None

        try:
            if "gsalr" in url.lower():
                event = scrape_gsalr(link)
            elif "craigslist" in url.lower():
                event = scrape_craigslist(link, geocoder)
            elif "pennlive" in url.lower():
                event = scrape_pennlive(link)
            elif "auctionzip" in url.lower():
                event = scrape_auctionzip(link)
        except Exception as exc:
            log(f"Failed scraping {link}: {exc}")

        if event:
            events.append(event)

        log(f"Visited: {link}")
        time.sleep(SCRAPE_SLEEP_SECONDS)

    return events


# =============================================================================
# PHOTO OCR
# =============================================================================

def cloud_ocr_image(image_path: Path) -> str:
    image_data = base64.b64encode(image_path.read_bytes()).decode()

    payload = {
        "apikey": OCR_SPACE_API_KEY,
        "base64Image": f"data:image/jpg;base64,{image_data}",
        "language": "eng",
        "isOverlayRequired": False,
        "scale": True,
        "OCREngine": 3,
    }

    response = requests.post(OCR_SPACE_URL, data=payload, timeout=60)
    response.raise_for_status()
    data = response.json()

    if data.get("IsErroredOnProcessing"):
        raise RuntimeError(str(data.get("ErrorMessage")))

    parsed = data.get("ParsedResults", [])
    return parsed[0].get("ParsedText", "").strip() if parsed else ""


def normalize_ocr_text(text: str) -> str:
    text = text.replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", text).strip()


def extract_photo_title(lines: List[str]) -> Optional[str]:
    keywords = [
        "yard sale", "garage sale", "community", "church", "parking lot sale",
        "estate sale", "moving sale", "bake sale", "flea market", "rummage sale",
    ]

    for line in lines:
        if any(k in line.lower() for k in keywords):
            return line.strip()

    return lines[0].strip() if lines else None


def extract_photo_address(text: str) -> Optional[str]:
    pattern = re.compile(
        r"\b\d{2,6}\s+"
        r"[A-Za-z0-9\s.'-]+?\s+"
        r"(?:Street|St|Road|Rd|Avenue|Ave|Court|Ct|Drive|Dr|Lane|Ln|"
        r"Boulevard|Blvd|Way|Circle|Cir|Place|Pl|Pike|Trail|Trl|Highway|Hwy)\b"
        r"(?:[, ]+\s*[A-Za-z\s.'-]+)?"
        r"(?:[, ]+\s*(?:PA|Pa|pa|Pennsylvania))?",
        re.IGNORECASE,
    )
    match = pattern.search(text)
    return match.group(0).strip(" ,") if match else None


def extract_photo_date(text: str) -> Optional[str]:
    patterns = [
        r"\b(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s*,?\s+"
        r"(?:January|February|March|April|May|June|July|August|September|October|November|December)"
        r"\s+\d{1,2}(?:st|nd|rd|th)?(?:,\s*\d{4})?",
        r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December)"
        r"\s+\d{1,2}(?:st|nd|rd|th)?(?:\s*&\s*\d{1,2})?(?:,\s*\d{4})?",
        r"\b\d{1,2}/\d{1,2}/\d{2,4}\b",
        r"\b\d{1,2}-\d{1,2}-\d{2,4}\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(0).strip()
    return None


def split_time(raw: str) -> Tuple[str, Optional[str]]:
    raw = raw.replace("–", "-").replace("—", "-")
    if re.search(r"\s+to\s+", raw, flags=re.IGNORECASE):
        parts = re.split(r"\s+to\s+", raw, maxsplit=1, flags=re.IGNORECASE)
    elif "-" in raw:
        parts = raw.split("-", 1)
    else:
        return raw.strip(), None

    return parts[0].strip(), parts[1].strip()


def extract_photo_time(text: str) -> Tuple[Optional[str], Optional[str]]:
    patterns = [
        r"\b\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)\s*(?:-|to)\s*(?:\?\?|\d{1,2}:?\d{0,2}\s*(?:AM|PM|am|pm)?)",
        r"\b\d{1,2}\s*(?:AM|PM|am|pm)\s*(?:-|to)\s*\d{1,2}\s*(?:AM|PM|am|pm)",
        r"\b\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)\b",
        r"\b\d{1,2}\s*(?:AM|PM|am|pm)\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return split_time(match.group(0).strip())
    return None, None


def parse_photo(image_path: Path) -> SaleEvent:
    raw_text = cloud_ocr_image(image_path)
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    full_text = normalize_ocr_text(raw_text)

    start_time, end_time = extract_photo_time(full_text)
    date_text = [x for x in [extract_photo_date(full_text), start_time, "-", end_time] if x]

    return SaleEvent(
        title=extract_photo_title(lines) or "Photo Yard Sale",
        description=raw_text,
        raw_address=extract_photo_address(full_text) or "",
        date_text=date_text or "No Dates Found",
        source_link=str(image_path),
        source="Cloud OCR",
        source_file=image_path.name,
    )


def process_photo_folder() -> List[SaleEvent]:
    events: List[SaleEvent] = []
    existing = load_json(PHOTO_OCR_JSON, [])
    processed_files = {sale.get("source_file") for sale in existing if sale.get("source_file")}

    if not PHOTO_FOLDER.exists():
        log(f"Photo folder does not exist: {PHOTO_FOLDER}")
        return events

    for image_path in PHOTO_FOLDER.iterdir():
        if not image_path.name.lower().endswith(IMAGE_EXTENSIONS):
            continue
        if image_path.name in processed_files:
            log(f"Skipping already processed photo: {image_path.name}")
            continue

        log(f"Reading photo: {image_path.name}")
        try:
            event = parse_photo(image_path)
            events.append(event)
            existing.append(asdict(event))
            save_json(PHOTO_OCR_JSON, existing)
            log(f"Photo found: {event.title} | {event.raw_address} | {event.date_text}")
        except Exception as exc:
            log(f"Failed to process photo {image_path.name}: {exc}")

    return events


# =============================================================================
# MAP HTML / JAVASCRIPT
# =============================================================================

def selection_panel_html() -> str:
    return r"""
<style>
  #saleSelectionPanel {
    position: fixed;
    top: 46%;
    right: -216px;
    width: 216px;
    transform: translateY(-50%);
    z-index: 999999;
    transition: right 0.35s ease;
    font-family: Arial, sans-serif;
  }
  #saleSelectionPanel.expanded { right: 0; }
  #saleSelectionTab {
    position: absolute;
    left: -38px;
    top: 50%;
    transform: translateY(-50%) rotate(-90deg);
    transform-origin: center;
    background: #1f1f1f;
    color: white;
    border: 1px solid #555;
    border-bottom: none;
    border-radius: 8px 8px 0 0;
    padding: 7px 12px;
    font-size: 12px;
    font-weight: bold;
    cursor: pointer;
    box-shadow: 0 2px 10px rgba(0,0,0,0.25);
    user-select: none;
    white-space: nowrap;
  }
  #saleSelectionBody {
    background: rgba(255,255,255,0.97);
    border: 1px solid #777;
    border-right: none;
    border-radius: 10px 0 0 10px;
    box-shadow: -2px 2px 10px rgba(0,0,0,0.25);
    padding: 12px;
  }
  #saleSelectionBody .sec-title {
    font-size: 13px;
    font-weight: bold;
    margin-bottom: 10px;
    text-align: center;
  }
  #saleSelectionBody button {
    display: block;
    width: 100%;
    margin: 8px 0;
    padding: 8px 10px;
    border: 1px solid #666;
    border-radius: 6px;
    background: #f5f5f5;
    cursor: pointer;
    font-weight: bold;
  }
  #saleSelectionBody button:hover { background: #e9e9e9; }
  #saleSelectionBody .count {
    margin-top: 8px;
    font-size: 12px;
    text-align: center;
  }
  .sale-select-wrap {
    margin-top: 6px;
    padding-top: 6px;
    border-top: 1px solid #ddd;
    font-family: Arial, sans-serif;
  }
  #dayFilterPanel {
    position: fixed;
    right: 8px;
    top: 8px;
    z-index: 999999;
    background: rgba(255,255,255,0.96);
    border: 1px solid #777;
    border-radius: 10px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.25);
    padding: 8px 10px;
    font-family: Arial, sans-serif;
    font-size: 13px;
  }
  #dayFilterPanel .filter-title {
    font-weight: bold;
    margin-bottom: 6px;
    text-align: center;
  }
  #dayFilterPanel label {
    display: block;
    margin: 4px 0;
    cursor: pointer;
    white-space: nowrap;
  }
</style>

<script>
window.saleData = window.saleData || {};
window.saleMarkers = window.saleMarkers || {};
window.saleDayFilterKey = "yard_sale_day_filters";
window.saleSelectionKey = "yard_sale_selected_ids";
window.saleSelectionPanelKey = "yard_sale_selection_panel_open";
window.saleStartKey = "yard_sale_start_id";

function getSelectedSaleIds() {
    try { return JSON.parse(localStorage.getItem(window.saleSelectionKey) || "[]"); }
    catch (err) { return []; }
}
function setSelectedSaleIds(ids) { localStorage.setItem(window.saleSelectionKey, JSON.stringify(ids)); }
function isSaleSelected(saleId) { return getSelectedSaleIds().includes(saleId); }
function toggleSaleSelection(saleId, isChecked) {
    const ids = new Set(getSelectedSaleIds());
    if (isChecked) { ids.add(saleId); } else { ids.delete(saleId); }
    setSelectedSaleIds(Array.from(ids));
    updateSelectedCount();
}
function getStartSaleId() { return localStorage.getItem(window.saleStartKey) || ""; }
function setStartSaleId(saleId) {
    if (saleId) { localStorage.setItem(window.saleStartKey, saleId); }
    else { localStorage.removeItem(window.saleStartKey); }
}
function isStartSale(saleId) { return getStartSaleId() === saleId; }
function toggleStartSale(saleId, isChecked) {
    if (isChecked) {
        setStartSaleId(saleId);
        const ids = new Set(getSelectedSaleIds());
        ids.add(saleId);
        setSelectedSaleIds(Array.from(ids));
    } else if (isStartSale(saleId)) {
        setStartSaleId("");
    }
    Object.keys(window.saleData).forEach(function(id) { syncSaleCheckbox(id); });
    updateSelectedCount();
}
function syncSaleCheckbox(saleId) {
    const checkbox = document.getElementById("sale-check-" + saleId);
    if (checkbox) { checkbox.checked = isSaleSelected(saleId); }
    const startCheckbox = document.getElementById("sale-start-" + saleId);
    if (startCheckbox) { startCheckbox.checked = isStartSale(saleId); }
}
function registerSaleData(saleObj) {
    window.saleData[saleObj.id] = saleObj;
    updateSelectedCount();
}
function registerSaleMarker(saleId, markerObj, mapObj) {
    window.saleMarkers[saleId] = { marker: markerObj, map: mapObj };
    applyDayFilters();
}
function registerSaleMarkerByName(saleId, markerName, mapName) {
    let tries = 0;
    function resolveMarker() {
        const markerObj = window[markerName];
        const mapObj = window[mapName];
        if (markerObj && mapObj) {
            registerSaleMarker(saleId, markerObj, mapObj);
            return;
        }
        tries += 1;
        if (tries < 80) {
            setTimeout(resolveMarker, 50);
        } else {
            console.warn("Could not register sale marker for day filter:", saleId, markerName, mapName);
        }
    }
    resolveMarker();
}
function getDayFilters() {
    try {
        const saved = localStorage.getItem(window.saleDayFilterKey);
        if (!saved) return { fri: true, sat: true };
        const parsed = JSON.parse(saved);
        return { fri: parsed.fri !== false, sat: parsed.sat !== false };
    } catch (err) {
        return { fri: true, sat: true };
    }
}
function setDayFilters(filters) {
    localStorage.setItem(window.saleDayFilterKey, JSON.stringify(filters));
}
function saleMatchesDay(dateText, dayKey) {
    const text = String(dateText || "").toLowerCase().replace(/<[^>]*>/g, " ").replace(/[,&;]+/g, " ");
    if (dayKey === "fri") { return /\b(fri|friday)\.?\b/.test(text); }
    if (dayKey === "sat") { return /\b(sat|saturday)\.?\b/.test(text); }
    return false;
}
function shouldShowSale(saleObj, filters) {
    const dateText = saleObj.date || saleObj.date_text || "";
    const matchesFri = saleMatchesDay(dateText, "fri");
    const matchesSat = saleMatchesDay(dateText, "sat");

    // OR behavior:
    // - Fri checked shows Friday sales, including Fri/Sat multi-day sales.
    // - Sat checked shows Saturday sales, including Fri/Sat multi-day sales.
    // - Both checked shows either Friday or Saturday sales.
    // - Neither checked hides Friday/Saturday sales.
    if (matchesFri || matchesSat) {
        return (filters.fri && matchesFri) || (filters.sat && matchesSat);
    }

    // Keep sales with no recognized Fri/Sat text visible so unknown dates do not vanish.
    return true;
}
function applyDayFilters() {
    const filters = getDayFilters();
    const friBox = document.getElementById("filter-fri");
    const satBox = document.getElementById("filter-sat");
    if (friBox) friBox.checked = filters.fri !== false;
    if (satBox) satBox.checked = filters.sat !== false;

    Object.keys(window.saleMarkers).forEach(function(saleId) {
        const entry = window.saleMarkers[saleId];
        const saleObj = window.saleData[saleId];
        if (!entry || !entry.marker || !entry.map || !saleObj) return;

        const showSale = shouldShowSale(saleObj, filters);
        const isOnMap = entry.map.hasLayer(entry.marker);

        if (showSale && !isOnMap) { entry.marker.addTo(entry.map); }
        if (!showSale && isOnMap) { entry.map.removeLayer(entry.marker); }
    });
}
function toggleDayFilter(dayKey, isChecked) {
    const filters = getDayFilters();
    filters[dayKey] = isChecked;
    setDayFilters(filters);
    applyDayFilters();
}
function updateSelectedCount() {
    const countEl = document.getElementById("selected-sale-count");
    if (!countEl) return;
    const ids = getSelectedSaleIds();
    countEl.textContent = ids.length + " sale" + (ids.length === 1 ? "" : "s") + " selected";
}
function clearSelectedSales() {
    setSelectedSaleIds([]);
    setStartSaleId("");
    Object.keys(window.saleData).forEach(function(saleId) { syncSaleCheckbox(saleId); });
    updateSelectedCount();
}
function exportSelectedSales() {
    const ids = getSelectedSaleIds();
    const startSaleId = getStartSaleId();
    const selected = ids
        .map(id => window.saleData[id])
        .filter(Boolean)
        .map(function(item) {
            return Object.assign({}, item, { status: item.id === startSaleId ? "Start" : "" });
        });

    const payload = {
        generated_at: new Date().toISOString(),
        selected_count: selected.length,
        start_sale_id: startSaleId,
        events: selected
    };

    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "selected_sales.json";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}
function setSalePanelExpanded(isExpanded) {
    const panel = document.getElementById("saleSelectionPanel");
    if (!panel) return;
    panel.classList.toggle("expanded", !!isExpanded);
    localStorage.setItem(window.saleSelectionPanelKey, isExpanded ? "true" : "false");
}
function toggleSalePanel() {
    const panel = document.getElementById("saleSelectionPanel");
    if (!panel) return;
    setSalePanelExpanded(!panel.classList.contains("expanded"));
}
document.addEventListener("DOMContentLoaded", function() {
    updateSelectedCount();
    const shouldOpen = localStorage.getItem(window.saleSelectionPanelKey) === "true";
    setSalePanelExpanded(shouldOpen);
    applyDayFilters();
});
</script>

<div id="saleSelectionPanel">
  <div id="saleSelectionTab" onclick="toggleSalePanel()">Selections</div>
  <div id="saleSelectionBody">
    <div class="sec-title">Sale Selection</div>
    <button onclick="exportSelectedSales()">Export Selected Sales</button>
    <button onclick="clearSelectedSales()">Clear Selected</button>
    <div id="selected-sale-count" class="count">0 sales selected</div>
  </div>
</div>

<div id="dayFilterPanel">
  <div class="filter-title">Show Sales<br><span style="font-size:11px; font-weight:normal;"></span></div>
  <label><input type="checkbox" id="filter-fri" checked onchange="toggleDayFilter('fri', this.checked)"> Fri.</label>
  <label><input type="checkbox" id="filter-sat" checked onchange="toggleDayFilter('sat', this.checked)"> Sat.</label>
</div>
"""


def sale_link_html(link: str) -> str:
    if link and link.lower().endswith(IMAGE_EXTENSIONS) and Path(link).exists():
        img_base64, mime = image_to_base64(link)
        if img_base64 and mime:
            image_src = f"data:{mime};base64,{img_base64}"
            return f'<a href="{image_src}" target="_blank" style="color: blue; font-weight: bold;">IMAGE OF SALE</a>'
        return "Image not available"

    if link:
        return f'<a href="{html.escape(link)}" target="_blank" style="color: blue">LINK TO SALE</a>'

    return ""


def build_popup_html(event: SaleEvent, sale_id: str) -> str:
    if address_contains_near(event.raw_address):
        query_addr = normalize_near_address(event.raw_address)
    else:
        query_addr = event.formatted_address or event.raw_address
    maps_web = event.google_maps_link or maps_search_url(query_addr)
    maps_geo = "geo:0,0?q=" + quote_plus(query_addr)

    maps_link_html = (
        f'<a href="{maps_geo}" target="_blank" style="color:#4da3ff; font-weight:bold;">'
        f'{html.escape(query_addr)}</a><br>'
        f'<a href="{maps_web}" target="_blank" style="color:#4da3ff;">(Open in browser)</a>'
    )

    details_html = "<br>".join(
        f'{html.escape(item["sheet"])} - {html.escape(item["title"])}: {html.escape(item["description"])}'
        for item in event.historical_details
    )

    return (
        f"<b>{html.escape(event.title)}</b><br>"
        f"{html.escape(event.description)}<br><br>"
        f"{maps_link_html}<br><br>"
        f"{event.formatted_date_html()}<br>"
        f"{sale_link_html(event.source_link)}<br><br>"
        f"{details_html}"
        f"<div class='sale-select-wrap'>"
        f"<label style='display:block; margin-bottom:6px;'>"
        f"<input type='checkbox' id='sale-check-{sale_id}' "
        f"onchange=\"toggleSaleSelection('{sale_id}', this.checked)\"> "
        f"Go to this sale"
        f"</label>"
        f"<label style='display:block;'>"
        f"<input type='checkbox' id='sale-start-{sale_id}' "
        f"onchange=\"toggleStartSale('{sale_id}', this.checked)\"> "
        f"Start here"
        f"</label>"
        f"</div>"
    )


def add_marker(mymap: folium.Map, event: SaleEvent, sale_id: str, popup_html: str) -> folium.Marker:
    title_lower = event.title.lower()
    location_lower = (event.formatted_address or event.raw_address).lower()
    link_lower = event.source_link.lower()

    if any(k in title_lower for k in HIGH_PRIORITY_NAME_KEYWORDS):
        icon = folium.Icon(color="red", icon="star")
    elif any(k in title_lower for k in MEDIUM_PRIORITY_KEYWORDS):
        icon = folium.Icon(color="blue", icon="star")
    elif any(k in location_lower for k in LOCATION_KEYWORDS):
        icon = folium.Icon(color="green", icon="home")
    elif any(k in link_lower for k in AUCTION_KEYWORDS):
        icon = folium.Icon(color="darkgreen", icon="gavel", prefix="fa")
    else:
        icon = folium.Icon(color="darkblue", icon="info-sign")

    marker = folium.Marker(
        location=[event.latitude, event.longitude],
        popup=folium.Popup(popup_html, max_width=300),
        icon=icon,
    )
    marker.add_to(mymap)
    return marker


def create_map(events: List[SaleEvent]) -> List[Dict[str, Any]]:
    log("Creating map...")

    mymap = folium.Map(
        location=MAP_CENTER,
        zoom_start=10,
        control_scale=True,
        control=False,
        tiles="OpenStreetMap",  # Default base layer
    )

    folium.TileLayer(
        tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        attr="Esri",
        name="Satellite",
        overlay=False,
        control=True,
        show=False,
    ).add_to(mymap)

    folium.TileLayer(
        tiles="https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png",
        name="CartoDB Dark Matter",
        attr="CartoDB",
        overlay=False,
        control=True,
        show=False,
    ).add_to(mymap)

    folium.TileLayer(
        tiles="https://services.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}",
        name="Esri Street Map",
        attr="Esri",
        overlay=False,
        control=True,
        show=False,
    ).add_to(mymap)

    mymap.add_child(folium.LayerControl(position="bottomleft"))
    mymap.get_root().html.add_child(Element(selection_panel_html()))

    export_events: List[Dict[str, Any]] = []
    unfound_addresses: List[str] = []

    for idx, event in enumerate(events, start=1):
        if event.title == "No Sale Data":
            continue

        event.historical_details = check_historical_sale_data(
            event.title,
            event.description,
            event.formatted_address or event.raw_address,
        )

        sale_id = f"sale_{idx}"

        if not event.has_location():
            unfound_addresses.append(
                f"<b>{html.escape(event.title)}</b><br>"
                f"{html.escape(event.raw_address)}<br>"
                f"{sale_link_html(event.source_link)}"
            )
            continue

        event_record = {
            "id": sale_id,
            "name": event.title,
            "description": event.description,
            "date": event.formatted_date_html(),
            "date_text": clean_spaces(str(event.date_text)),
            "address": event.formatted_address or event.raw_address,
            "latitude": event.latitude,
            "longitude": event.longitude,
            "source": event.source,
            "source_link": event.source_link,
            "details": event.historical_details,
            "google_maps_link": event.google_maps_link or maps_search_url(event.formatted_address or event.raw_address),
        }
        export_events.append(event_record)

        popup = build_popup_html(event, sale_id)
        marker = add_marker(mymap, event, sale_id, popup)

        sale_js_obj = json.dumps(event_record, ensure_ascii=False).replace("</", "<\\/")
        register_js = f"""
<script>
registerSaleData({sale_js_obj});
registerSaleMarkerByName("{sale_id}", "{marker.get_name()}", "{mymap.get_name()}");
{marker.get_name()}.on('popupopen', function() {{
    setTimeout(function() {{
        syncSaleCheckbox("{sale_id}");
    }}, 40);
}});
</script>
"""
        mymap.get_root().html.add_child(Element(register_js))

    if unfound_addresses:
        dropdown_html = (
            '<div style="position: absolute; top: 10px; left: 50px; right: 10px; z-index:2000; '
            'background-color: black; color: white; padding: 10px; border-radius: 5px; '
            'box-shadow: 2px 2px 5px rgba(0,0,0,0.3);">'
            ' <label for="missingLocations"><b>Unfound Addresses:</b></label> '
            '<select id="missingLocations" onchange="openLink(this)" style="width:100%;">'
            '<option value="">-- Select an Address --</option>'
        )
        for addr in unfound_addresses:
            match = re.search(r'href="([^"]*)"', addr)
            unfound_link = match.group(1) if match else "#"
            option_text = re.sub(r"<[^>]+>", " ", addr)
            option_text = re.sub(r"\s+", " ", option_text).strip()
            dropdown_html += f'<option value="{html.escape(unfound_link)}">{html.escape(option_text)}</option>'
        dropdown_html += (
            '</select></div>'
            '<script>function openLink(select){var url=select.value; if(url && url !== "#"){window.open(url, "_blank");}}</script>'
        )
        mymap.get_root().html.add_child(Element(dropdown_html))

    updated = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    header_html = (
        f'<div style="position: fixed; bottom: 5px; left: 50%; transform: translateX(-50%); '
        f'z-index: 9999; background-color: white; padding: 5px 10px; border-radius: 2px; '
        f'font-size: 10px; box-shadow: 0 2px 6px rgba(0,0,0,0.3);">'
        f'Yard Sale Finder v{VERSION} | Updated: {updated}</div>'
    )
    mymap.get_root().html.add_child(Element(header_html))

    mymap.save(str(MAP_FILE))
    log(f"Map saved as {MAP_FILE}")

    return export_events


# =============================================================================
# MAIN
# =============================================================================

def main() -> None:
    log(f"Running Yard Sale Finder v{VERSION}")
    geocoder = GoogleGeocoder(GOOGLE_MAPS_API_KEY)

    all_events: List[SaleEvent] = []

    for url in build_main_urls():
        all_events.extend(scrape_site(url, geocoder))

    all_events.extend(process_photo_folder())

    # Remove obvious duplicates by title+address/source link.
    deduped: List[SaleEvent] = []
    seen = set()
    for event in all_events:
        key = (
            clean_spaces(event.title).lower(),
            clean_spaces(event.raw_address or event.formatted_address).lower(),
            clean_spaces(event.source_link).lower(),
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(event)

    log(f"Found {len(deduped)} unique sale events before geocoding.")

    geocode_events(deduped, geocoder)

    export_events = create_map(deduped)

    export_data = {
        "version": VERSION,
        "generated_at": datetime.now().isoformat(),
        "event_count": len(export_events),
        "events": export_events,
    }
    save_json(EXPORT_JSON_FILE, export_data)
    log(f"Exported {len(export_events)} events to {EXPORT_JSON_FILE}")


if __name__ == "__main__":
    main()
