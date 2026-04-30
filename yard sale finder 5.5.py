import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import pyap
from datetime import datetime, timedelta
import folium
from geopy.geocoders import Nominatim
import time
import pgeocode
import os
import subprocess
import re
import pandas as pd
from rapidfuzz import fuzz
import json
import random
from urllib.parse import quote_plus
export_events = []
import base64


PHOTO_FOLDER = "/storage/emulated/0/Download"
OUTPUT_JSON = "/storage/emulated/0/Download/yard_sales_from_photos.json"

OCR_SPACE_API_KEY = "helloworld"  # free demo key, limited. Better to get your own.
OCR_SPACE_URL = "https://api.ocr.space/parse/image"

IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".webp")


def check_historical_sale_data(name, description, address):
    RED = '\033[91m'
    GREEN = '\033[92m'
    RESET = '\033[0m'

    history = []
    full_text = (name + " " + description + " " + address).lower()
    excel_path = 'yardsale_descriptions.xlsx'
    xls = pd.ExcelFile(excel_path)
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        if df.shape[0] >= 2:
            titles = df.iloc[:, 0].fillna("").astype(str)
            descriptions = df.iloc[:, 1].fillna("").astype(str)
            
            #print(full_text)
            
            for title, descriptions in zip(titles, descriptions):
                #print(title.lower())
                #print(f"Chexking: {title.lower()} vs. {full_text.lower()}")
                #print(fuzz.partial_ratio(title.lower(), full_text.lower()))
                
                #if title.lower() in full_text:
                if fuzz.partial_ratio(title.lower(), full_text.lower()) >= 87:
                    print(f"Found: {GREEN}{title}{RESET} in {RED}{full_text}{RESET}")
                    history.append([sheet_name, title.strip(), descriptions.strip()])
                    
    for item,item2,item3 in history:
    	print(item, item2, item3)
    return history


# Function to get all relevant links from the main URL

def get_relevant_pennlive_links(url, keywords, links_to_remove):
    keywords=["Show more"]
	
    headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"Request failed: {e}")
        return []
     
    print(keywords)
    soup = BeautifulSoup(response.content, 'html.parser')
    links = soup.find_all('a')
    relevant_links = []
    
    
    for link in links:
        href = link.get('href')
        if href:
            full_url = urljoin(url, href)
            link_text = link.get_text()
            if any(keyword.lower() in link_text.lower() for keyword in keywords):
                if not any(remove.lower() in link_text.lower() for remove in links_to_remove):
                    print(link_text)
                    relevant_links.append((link_text, full_url))
                                
    return relevant_links
    
    
def get_relevant_auctionzip_links(url):
    import random
    headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/128.0.0.0 Safari/537.36"
}


    tries = 0

    while tries < 10:
    	r = requests.get(url, headers=headers)
    	soup = BeautifulSoup(r.text, "html.parser")
    	#x = random.uniform(1,5.5)
    	#time.sleep(x)
    	print(r.status_code)
    	if r.status_code == 200:
    		tries = 10
    		print("End")
    	tries = tries + 1
    	print("try number ", tries)
        
    relevant_links = []

# Find all listing containers
    listings = soup.find_all("div", class_="az-ListOfLlisting-body")

    for listing in listings:
        # find the main link
        a = listing.find("a", class_="az-ListOfLlisting-body__link")
        if a and a.get("href"):
            href = a["href"]
        # normalize markdown-style links [url](url)
            if "(" in href and ")" in href:
                href = href[href.find("(") + 1:href.find(")")]
            if href.startswith("//"):
                href = "https:" + href
            elif href.startswith("/"):
                href = "https://www.auctionzip.com" + href
            title = a.get_text(strip=True).replace("\xa0", " ")
            print("🔗", href)
            print("🪧", title)
            print("-" * 70)
            relevant_links.append((href, href))
   
    return relevant_links
	
	
def get_relevant_links(url, keywords, links_to_remove):
    
    headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
}

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"Request failed: {e}")
        return []

    soup = BeautifulSoup(response.content, 'html.parser')
    links = soup.find_all('a')
    relevant_links = []
    
    for link in links:
        href = link.get('href')
        if href:
            full_url = urljoin(url, href)
            link_text = link.get_text()
            if any(keyword.lower() in link_text.lower() for keyword in keywords):
                if not any(remove.lower() in link_text.lower() for remove in links_to_remove):
                    relevant_links.append((link_text, full_url))
                    
    return relevant_links

# Function to scrape Craigslist sales
def scrape_and_print_craigslist(url, addresses, names, descriptions, dates):
    print("SCRAPING CRAIGSLIST...")
    try:
        response = requests.get(url)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"Request failed: {e}")
        return addresses, names, descriptions

    soup = BeautifulSoup(response.content, 'html.parser')
    # Find all meta tags
    meta_tags = soup.find_all('meta')
    for tag in meta_tags:
    	attributes = {k.lower().strip(): v for k, v in tag.attrs.items()}
    	content = attributes.get('content', '')
    	
    	if attributes.get('property') in ['og:title', 'og:description']:
            print(f" {content}")
    	if attributes.get("property") in ["og:title"]:
        	name = content
        	names.append(name)
    	if attributes.get("property") in ["og:description"]:
        	description = content
        	descriptions.append(description)
        	geo_position = soup.find("meta", {"name": "geo.position"})["content"]
        	latitude, longitude = map(float, geo_position.split(";"))
        	latitude = round(latitude, 4)
        	longitude = round(longitude, 4)
        	geolocator = Nominatim(user_agent="zip_lookup")
        	location = geolocator.reverse((latitude, longitude), exactly_one=True)
        	zipcode = location.raw.get("address", {}).get("postcode", "ZIP not found")
        	#print(f"Coordinates: ({latitude}, {longitude})")
        	print(f"ZIP Code: {zipcode}")
        	       	
    	name_attr = attributes.get('name', '').strip()  # Strip extra spaces

    	if name_attr in ['geo.region']:
    	   country = content[:2]
    	   state = content[-2:]
    	if name_attr in ['geo.position']:
    	   position = content
    	if name_attr in ['geo.placename']:
    	   city = content        
    
    address = soup.find('div', class_='mapaddress').text
    address = f"{address}, {city}, {state} {zipcode}"
    
    if address:
    	print(f'Address: {address}\n')
    	addresses.append(address)
    	
    
    sale_dates = [date.text.strip() for date in soup.find_all("a", class_="valu")]
    start_time = soup.find("div", class_="attr sale_time").find("span", class_="valu").text.strip()
    if sale_dates:
    	print(sale_dates, start_time)
    	sale_dates.append(start_time)
    else:
    	sale_dates = "No Dates Found"
    dates.append(sale_dates)
     
    return addresses, names, descriptions, dates






def cloud_ocr_image(image_path):
    with open(image_path, "rb") as f:
        image_data = base64.b64encode(f.read()).decode()

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
        raise Exception(data.get("ErrorMessage"))

    parsed = data.get("ParsedResults", [])
    if not parsed:
        return ""

    return parsed[0].get("ParsedText", "").strip()


def normalize_text(text):
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def extract_title(lines):
    keywords = [
        "yard sale", "garage sale", "community", "church",
        "parking lot sale", "estate sale", "moving sale",
        "bake sale", "flea market"
    ]

    for line in lines:
        low = line.lower()
        if any(k in low for k in keywords):
            return line.strip()

    return lines[0].strip() if lines else None


def extract_address(text):
    pattern = re.compile(
        r"\b\d{2,6}\s+"
        r"[A-Za-z0-9\s.'-]+?\s+"
        r"(?:Street|St|Road|Rd|Avenue|Ave|Court|Ct|Drive|Dr|Lane|Ln|"
        r"Boulevard|Blvd|Way|Circle|Cir|Place|Pl|Pike|Trail|Trl|Highway|Hwy)\b"
        r"(?:[, ]+\s*[A-Za-z\s.'-]+)?"
        r"(?:[, ]+\s*(?:PA|Pa|pa|Pennsylvania))?",
        re.IGNORECASE
    )

    match = pattern.search(text)
    return match.group(0).strip(" ,") if match else None


def extract_date(text):
    patterns = [
        r"\b(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s*,?\s+"
        r"(?:January|February|March|April|May|June|July|August|September|October|November|December)"
        r"\s+\d{1,2}(?:st|nd|rd|th)?(?:,\s*\d{4})?",

        r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December)"
        r"\s+\d{1,2}(?:st|nd|rd|th)?(?:\s*&\s*\d{1,2})?(?:,\s*\d{4})?",

        r"\b\d{1,2}/\d{1,2}/\d{2,4}\b",
        r"\b\d{1,2}-\d{1,2}-\d{2,4}\b",
    ]

    for p in patterns:
        match = re.search(p, text, re.IGNORECASE)
        if match:
            return match.group(0).strip()

    return None


def split_time(raw):
    raw = raw.replace("–", "-").replace("—", "-")

    if " to " in raw.lower():
        parts = re.split(r"\s+to\s+", raw, flags=re.IGNORECASE)
    elif "-" in raw:
        parts = raw.split("-", 1)
    else:
        return raw.strip(), None

    return parts[0].strip(), parts[1].strip()


def extract_time(text):
    patterns = [
        r"\b\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)\s*(?:-|to)\s*(?:\?\?|\d{1,2}:?\d{0,2}\s*(?:AM|PM|am|pm)?)",
        r"\b\d{1,2}\s*(?:AM|PM|am|pm)\s*(?:-|to)\s*\d{1,2}\s*(?:AM|PM|am|pm)",
        r"\b\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)\b",
        r"\b\d{1,2}\s*(?:AM|PM|am|pm)\b",
    ]

    for p in patterns:
        match = re.search(p, text)
        if match:
            return split_time(match.group(0).strip())

    return None, None


def parse_photo(image_path):
    raw_text = cloud_ocr_image(image_path)

    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    full_text = normalize_text(raw_text)

    start_time, end_time = extract_time(full_text)

    return {
        "title": extract_title(lines),
        "address": extract_address(full_text),
        "date": extract_date(full_text),
        "start_time": start_time,
        "end_time": end_time,
        "description": raw_text,
        "source": "cloud_ocr",
        "source_file": os.path.basename(image_path)
    }


def load_existing():
    if not os.path.exists(OUTPUT_JSON):
        return []

    try:
        with open(OUTPUT_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def image_to_base64(path):
    try:
        ext = os.path.splitext(path)[1].lower()

        if ext in [".jpg", ".jpeg"]:
            mime = "image/jpeg"
        elif ext == ".png":
            mime = "image/png"
        elif ext == ".webp":
            mime = "image/webp"
        else:
            mime = "image/jpeg"

        with open(path, "rb") as img:
            encoded = base64.b64encode(img.read()).decode()

        return encoded, mime

    except Exception as e:
        print(f"Could not encode image {path}: {e}")
        return None, None
        

def save_sales(sales):
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(sales, f, indent=4, ensure_ascii=False)


def process_folder(addresses, names, descriptions, dates, url_list):
    sales = load_existing()

    processed_files = {
        sale.get("source_file")
        for sale in sales
        if sale.get("source_file")
    }

    for filename in os.listdir(PHOTO_FOLDER):
        if not filename.lower().endswith(IMAGE_EXTENSIONS):
            continue

        if filename in processed_files:
            print(f"Skipping already processed: {filename}")
            continue

        image_path = os.path.join(PHOTO_FOLDER, filename)

        print(f"Reading photo: {filename}")

        try:
            sale = parse_photo(image_path)
            sales.append(sale)

            print("Found:")
            print(f"  Title: {sale['title']}")
            names.append(sale['title'])
            print(f"  Address: {sale['address']}")
            addresses.append(sale['address'])
            print(f"  Date: {sale['date']}")
            dates.append([sale['date'], " ", sale['start_time'], " - ", sale['end_time']])
            print(f"  Time: {sale['start_time']} - {sale['end_time']}")
            print(f"  Description: {sale['description']}")
            descriptions.append(sale['description'])
            print()
            url_list.append(image_path)

        except Exception as e:
            print(f"Failed to process {filename}: {e}")

    #save_sales(sales)
    #print(f"Saved {len(sales)} total sales to {OUTPUT_JSON}")
    print(url_list)    
        
    return addresses, names, descriptions, dates, url_list 


def scrape_and_print_auctionzip(url,addresses, names, descriptions, dates):
    
    #url="https://www.auctionzip.com/Listings/4080535.html?kwd=&zip=17055&category=0"

# Browser-like headers
    headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/128.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,"
        "application/signed-exchange;v=b3;q=0.7"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Referer": "https://www.auctionzip.com/",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-User": "?1",
    "Sec-Fetch-Dest": "document",
    "DNT": "1"
}

# --- SETUP SESSION ---
    session = requests.Session()
    session.headers.update(headers)

# --- FETCH PAGE WITH RETRIES ---
    def fetch_with_retries(url, retries=5):
        for attempt in range(1, retries + 1):
            try:
                response = session.get(url, timeout=10)
            
            # Handle rate limiting or forbidden
                if response.status_code != 200:
                    wait = random.uniform(3, 7)
                    print(f"[{response.status_code}] Blocked, waiting {wait:.1f}s then retrying ({attempt}/{retries})...")
                    time.sleep(wait)
                    continue
            
                response.raise_for_status()
                return response

            except requests.RequestException as e:
                wait = random.uniform(2, 6)
                print(f"Error: {e} — waiting {wait:.1f}s (attempt {attempt}/{retries})...")
                time.sleep(wait)
    
        raise Exception("Failed to fetch page after multiple attempts.")

# --- RUN ---
    r = fetch_with_retries(url)
    print(f"Success! Status code: {r.status_code}")

    # --- PARSE PAGE ---
    soup = BeautifulSoup(r.text, "html.parser")


    # Find the <script> tag containing 'dataLayer.push'
    script_tag = soup.find("script", string=re.compile(r"dataLayer\.push"))

    # Extract the JSON-like part inside the dataLayer.push(...)
    match = re.search(r"dataLayer\.push\((\{.*?\})\);", script_tag.string, re.DOTALL)

    if match:
        json_text = match.group(1)
    # Clean up to valid JSON
        #json_text = json_text.replace("'", '"')
        data = json.loads(json_text)
   
    
        seller = data.get("sellerName")
        sale_date = data.get("SaleDate")
        sale_date = sale_date.replace(",","")
        auction_name = data.get("name")
        auction_name = auction_name.replace(",", "")
        location = data.get("sellerLocation")
        location = location.replace(",","")
        description = soup.find("meta", attrs={"name": "description"}).get("content", "")
        description = description.replace("/","")

        print(f"Seller: {seller}")        
        print(f"Sale Date: {sale_date}")
        dates.append(sale_date)
        print(f"Auction Name: {auction_name}")
        names.append(auction_name)
        print(f"Description: {description}")
        descriptions.append(description)
        print()
        print(f"Location: {location}")
        addresses.append(location)

    else:
        print("Could not find dataLayer JSON.")
        
        
    return addresses, names, descriptions, dates    
    
    
    

def scrape_and_print_gsalr(url, addresses, names, descriptions, dates):
    print("SCRAPING GSALR...")
    try:
        response = requests.get(url)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"Request failed: {e}")
        return addresses, names, descriptions, dates

    soup = BeautifulSoup(response.content, 'html.parser')

    name = soup.find(itemprop='name').text if soup.find(itemprop='name') else "No Sale Data"
    names.append(name)
    
    description = soup.find(itemprop='description').text if soup.find(itemprop='description') else "No Sale Data"
    descriptions.append(description)
    
    sale_times = []
    
    for time_div in soup.find_all("div", class_="sale-date-cards"):
    	text = time_div.text.strip()
    	print(text)
    	result = re.sub(r"(PM)", r"\1\n", text)
    	print(result)
    	if result == "":
    		result = "No Times Found"
    	sale_times.append(result)
    	#sale_times.append(time_div.text.strip())
    	
    if sale_times != "":
    	dates.append(sale_times)
    else:
        dates.append("No Times Found")
        
    street = soup.find(itemprop='streetAddress').text if soup.find(itemprop='streetAddress') else "No Sale Data"
    city = soup.find(itemprop="addressLocality").text if soup.find(itemprop="addressLocality") else "No Sale Data"
    state = soup.find(itemprop="addressRegion").text if soup.find(itemprop="addressRegion") else "No Sale Data"
    zipcode = soup.find(itemprop="postalCode").text if soup.find(itemprop="postalCode") else "No Sale Data"
    
    address = f"{street}, {city}, {state} {zipcode}"
    addresses.append(address)

    print(f"Title: {name}\n")
    print(f"Description: {description}\n")
    print(f"Address: {address}\n")
    print("Sale Times:")
    for time in sale_times:
    	print(time)

    return addresses, names, descriptions, dates 
    


    
def scrape_and_print_pennlive(url, addresses, names, descriptions, dates):
    print("SCRAPING PENNLIVE...")
    headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
}

    try:
        response = requests.get(url, headers = headers)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"Request failed: {e}")
        return addresses, names, descriptions


    soup = BeautifulSoup(response.content, 'html.parser')     
      

    # Locate the main ad block using consistent class
    main_ad = soup.find("div", class_="sr_ad_frame")
    if not main_ad:
        print("Main ad block not found.")
        return None

    
        # Correct title extraction path
    title_tag = main_ad.find("span", class_="sr_ad_title")
    title = title_tag.get_text(strip=True) if title_tag else "N/A"

    # Extract description from <p itemprop="description">
    desc_tag = main_ad.find("p", itemprop="description")
    description = desc_tag.get_text(strip=True) if desc_tag else "N/A"
    
    text = description
    date_match = re.search(r'\b(Saturday),?\s+([A-Za-z]+ \d{1,2})', text, re.IGNORECASE)
    day = f"{date_match.group(1).capitalize()}, {date_match.group(2)}" if date_match else "N/A"

    # Time (e.g., 7 - 2, 7am - 2pm)
    time_match = re.search(r'\b\d{1,2}\s*[-to]+\s*\d{1,2}\b', text)
    time = time_match.group().replace("to", "-") if time_match else "N/A"
    
    street_match = re.search(r'\d{3,5}\s+[\w\s]+(?:St|Ave|Rd|Dr|Ln|Blvd)\.?', text)
    street = street_match.group().strip() if street_match else "N/A"

    # City/Zip
    city_match = re.search(r'New Cumberland,?\s*(\d{5})?', text)
    city = city_match.group().strip() if city_match else "N/A"
    


    date = [day, time]
    address = (street + city)
    print (f"Title: {title}")
    print(f"Description: {description}")
    print(f"Time: {date}")
    print(address)
    
    addresses.append(address)
    names.append(title)
    descriptions.append(description)
    dates.append(time)   
    
     
    return addresses, names, descriptions, dates
    
    
    
def get_date():
	from datetime import datetime
	import pytz
	
	# Define Eastern Time Zone
	eastern = pytz.timezone('America/New_York')
	
	# Get current time in EST
	current_time_est = datetime.now(eastern)
	
	# Format the date and time as needed
	formatted_time = current_time_est.strftime("%d-%m-%Y %H:%M:%S")
	
	print("Current Date and Time in EST:", formatted_time)		
	return formatted_time
	
	
# Function to generate a map with sale locations
def create_map(addresses, names, descriptions, dates, url_list):
    print("Creating map...")
    map_center = [40.2379, -76.9223]  # Default center

    mymap = folium.Map(
        location=map_center,
        zoom_start=10,
        control_scale=True,
        control=False,
        tiles='OpenStreetMap',
    )

    folium.TileLayer(
        tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
        attr='Esri',
        name='Satellite',
        overlay=False,
        control=True
    ).add_to(mymap)

    folium.TileLayer(
        tiles='https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png',
        name='CartoDB Dark Matter',
        attr='Carto'
    ).add_to(mymap)

    folium.TileLayer(
        tiles='https://services.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}',
        name='Esri Street Map',
        attr='Esri'
    ).add_to(mymap)

    mymap.add_child(folium.LayerControl(position="bottomleft"))

    GOOGLE_API_KEY = "AIzaSyBjDVzoi-JgNUaBKyRgMhxSqipgBrxWkG8"
    geolocator = Nominatim(user_agent="geo_lookup")

    unfound_addresses = []
    history = ["N", "N", "N"]

    from folium import Element

    selection_ui = r"""
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

  #saleSelectionPanel.expanded {
    right: 0;
  }

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

  #saleSelectionBody button:hover {
    background: #e9e9e9;
  }

  #saleSelectionBody .count {
    margin-top: 8px;
    font-size: 12px;
    text-align: center;
  }

  .sale-select-wrap {
    margin-top: 8px;
    padding-top: 8px;
    border-top: 1px solid #ddd;
    font-family: Arial, sans-serif;
  }
</style>

<script>
window.saleData = window.saleData || {};
window.saleSelectionKey = "yard_sale_selected_ids";
window.saleSelectionPanelKey = "yard_sale_selection_panel_open";
window.saleStartKey = "yard_sale_start_id";

function getSelectedSaleIds() {
    try {
        return JSON.parse(localStorage.getItem(window.saleSelectionKey) || "[]");
    } catch (err) {
        return [];
    }
}

function setSelectedSaleIds(ids) {
    localStorage.setItem(window.saleSelectionKey, JSON.stringify(ids));
}

function isSaleSelected(saleId) {
    return getSelectedSaleIds().includes(saleId);
}

function toggleSaleSelection(saleId, isChecked) {
    const ids = new Set(getSelectedSaleIds());
    if (isChecked) {
        ids.add(saleId);
    } else {
        ids.delete(saleId);
    }
    setSelectedSaleIds(Array.from(ids));
    updateSelectedCount();
}

function getStartSaleId() {
    return localStorage.getItem(window.saleStartKey) || "";
}

function setStartSaleId(saleId) {
    if (saleId) {
        localStorage.setItem(window.saleStartKey, saleId);
    } else {
        localStorage.removeItem(window.saleStartKey);
    }
}

function isStartSale(saleId) {
    return getStartSaleId() === saleId;
}

function toggleStartSale(saleId, isChecked) {
    if (isChecked) {
        setStartSaleId(saleId);
        const ids = new Set(getSelectedSaleIds());
        ids.add(saleId);
        setSelectedSaleIds(Array.from(ids));
    } else if (isStartSale(saleId)) {
        setStartSaleId("");
    }

    Object.keys(window.saleData).forEach(function(id) {
        syncSaleCheckbox(id);
    });
    updateSelectedCount();
}

function syncSaleCheckbox(saleId) {
    const checkbox = document.getElementById("sale-check-" + saleId);
    if (checkbox) {
        checkbox.checked = isSaleSelected(saleId);
    }

    const startCheckbox = document.getElementById("sale-start-" + saleId);
    if (startCheckbox) {
        startCheckbox.checked = isStartSale(saleId);
    }
}

function registerSaleData(saleObj) {
    window.saleData[saleObj.id] = saleObj;
    updateSelectedCount();
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
    Object.keys(window.saleData).forEach(function(saleId) {
        syncSaleCheckbox(saleId);
    });
    updateSelectedCount();
}

function exportSelectedSales() {
    const ids = getSelectedSaleIds();
    const startSaleId = getStartSaleId();
    const selected = ids
        .map(id => window.saleData[id])
        .filter(Boolean)
        .map(function(item) {
            return Object.assign({}, item, {
                status: item.id === startSaleId ? "Start" : ""
            });
        });

    const payload = {
        generated_at: new Date().toISOString(),
        selected_count: selected.length,
        start_sale_id: startSaleId,
        events: selected
    };

    const blob = new Blob(
        [JSON.stringify(payload, null, 2)],
        { type: "application/json" }
    );

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
"""
    mymap.get_root().html.add_child(Element(selection_ui))

    for idx, (name, description, address, date, link) in enumerate(zip(names, descriptions, addresses, dates, url_list), start=1):
        if name == "No Sale Data":
            continue

        history = check_historical_sale_data(name, description, address)

        details = "<br>".join(
            [f"{item} - {item2}: {item3}" for item, item2, item3 in history if len(history) >= 1]
        )

        print("*" * 15)

        try:
            location = geolocator.geocode(address, timeout=10)
        except Exception:
            print(f"Timeout error with {address} in neomatium geolocator")
            location = None

        if not location:
            google_url = f"https://maps.googleapis.com/maps/api/geocode/json?address={address}&key={GOOGLE_API_KEY}"
            try:
                response = requests.get(google_url, timeout=10).json()
                if response["status"] == "OK":
                    location_data = response["results"][0]["geometry"]["location"]
                    latitude, longitude = location_data["lat"], location_data["lng"]
                    if not 'near' in address:
                    	formatted_address = response["results"][0]["formatted_address"]
                    else:
                    	formatted_address = address
                    location = type(
                        "Location",
                        (),
                        {"latitude": latitude, "longitude": longitude, "address": formatted_address},
                    )
            except requests.RequestException as e:
                print(f"Request failed: {e}")

        if not location:
            location = type(
                "Location",
                (),
                {"latitude": "40.3582913", "longitude": "-76.9298062", "address": "108 Banbury circle, hummelstown pa 17036"},
            )

        if isinstance(date, (list, tuple)):
            formatted_date = "<br>".join(date)
        else:
            formatted_date = date

        location_to_check = location.address
        query_addr = (location.address if location and getattr(location, "address", None) else address).strip()
        encoded_addr = quote_plus(query_addr)

        maps_web = "https://www.google.com/maps/search/?api=1&query=" + encoded_addr
        maps_geo = "geo:0,0?q=" + encoded_addr

        maps_link_html = (
            '<a href="' + maps_geo + '" target="_blank" '
            'style="color:#4da3ff; font-weight:bold;">'
            + query_addr
            + '</a><br>'
            '<a href="' + maps_web + '" target="_blank" '
            'style="color:#4da3ff;">(Open in browser)</a>'
        )
        print(link)
        if link and link.lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
        	print("Part A works")
        	img_base64, mime = image_to_base64(link)

        	if img_base64:
        		print("Part B works")
        		sale_link = f'''
        <a href="#" onclick="
            var w = window.open('');
            w.document.write('<html><body style=\\'margin:0;background:#111;\\'><img src=\\'{mime};base64,{img_base64}\\' style=\\'width:100%;height:auto;\\'></body></html>');
            w.document.title = 'Sale Image';
        " style="color: blue; font-weight: bold;">
        IMAGE OF SALE
        </a>
        '''
        	else:
        		sale_link = "Image not available"
        else:
        	sale_link = f'<a href="{link}" target="_blank" style="color: blue">LINK TO SALE</a>'

        if location:
            keywords_to_check = ["community", "neighborhood", "multi", "flea", "development"]
            keywords_to_check2 = ["school", "church", "center", "building", "YMCA"]
            auction_keywords = ["auction", "bidding"]
            print(f"\nFound: {location.address}")

            sale_id = f"sale_{idx}"

            event_record = {
                "id": sale_id,
                "name": name,
                "description": description,
                "date": formatted_date,
                "address": location.address,
                "latitude": location.latitude,
                "longitude": location.longitude,
                "source_link": link,
                "details": details,
                "google_maps_link": maps_web,
            }

            export_events.append(event_record)

            popup_html = (
                f"<b>{name}</b><br>"
                f"{description}<br><br>"
                f"{maps_link_html}<br><br>"
                f"{formatted_date}<br>"
                f"{sale_link}<br><br>"
                f"{details}"
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

            if any(var in name.lower() for var in keywords_to_check):
                marker = folium.Marker(
                    location=[location.latitude, location.longitude],
                    popup=folium.Popup(popup_html, max_width=300),
                    icon=folium.Icon(color="red", icon="star"),
                )
            elif any(var in location_to_check.lower() for var in keywords_to_check2):
                marker = folium.Marker(
                    location=[location.latitude, location.longitude],
                    popup=folium.Popup(popup_html, max_width=300),
                    icon=folium.Icon(color="green", icon="home"),
                )
            elif any(var in link.lower() for var in auction_keywords):
                marker = folium.Marker(
                    location=[location.latitude, location.longitude],
                    popup=folium.Popup(popup_html, max_width=300),
                    icon=folium.Icon(color="darkgreen", icon="gavel", prefix="fa"),
                )
            else:
                marker = folium.Marker(
                    location=[location.latitude, location.longitude],
                    popup=folium.Popup(popup_html, max_width=300),
                    icon=folium.Icon(color="darkblue", icon="info-sign"),
                )

            marker.add_to(mymap)

            sale_js_obj = json.dumps(event_record).replace("</", "<\\/")
            register_js = f"""
<script>
registerSaleData({sale_js_obj});
{marker.get_name()}.on('popupopen', function() {{
    setTimeout(function() {{
        syncSaleCheckbox("{sale_id}");
    }}, 40);
}});
</script>
"""
            mymap.get_root().html.add_child(Element(register_js))
            time.sleep(2.5)
        else:
            print(f"\n{address} NOT FOUND")
            unfound_addresses.append(f"<b>{name}</b><br>{address}<br>{sale_link}<br><br>{details}")

    if unfound_addresses:
        dropdown_html = (
            '<div style="position: absolute; top: 10px; left: 50px; right: 10px; z-index:2000; background-color: black; padding: 10px; border-radius: 5px; box-shadow: 2px 2px 5px rgba(0,0,0,0.3);">'
            ' <label for="missingLocations"><b>Unfound Addresses:</b></label> '
            '<select id="missingLocations" onchange="openLink(this)" style="width:100%;">'
            '<option value="">-- Select an Address --</option>'
        )
        for addr in unfound_addresses:
            match = re.search(r'href="([^"]*)"', addr)
            unfound_link = match.group(1) if match else "#"
            option_text = re.sub(r"<[^>]+>", " ", addr)
            option_text = re.sub(r"\s+", " ", option_text).strip()
            dropdown_html += f'<option value="{unfound_link}">{option_text}</option>'
        dropdown_html += (
            '</select></div>'
            '<script>function openLink(select){var url=select.value; if(url && url !== "#"){window.open(url, "_blank");}}</script>'
        )
    else:
        dropdown_html = ""

    dropdown_element = Element(dropdown_html)
    mymap.get_root().html.add_child(dropdown_element)

    updatetime = get_date()
    Header_Text = f"Updated: {updatetime}"
    Header_html = f"""<div style="position: fixed; bottom: 5px; left: 50%; transform: translateX(-50%); z-index: 9999; background-color: white; padding: 5px 10px; border-radius: 2px; font-size: 10px; box-shadow: 0 2px 6px rgba(0,0,0,0.3);"> {Header_Text} </div> """
    header_element = Element(Header_html)
    mymap.get_root().html.add_child(header_element)


    mymap.save("map.html")
    print("Map saved as map.html")


# Main function to handle scraping and mapping
def main_scrape(url, keywords, addresses, descriptions, names, dates, url_list):
    
    # Exclude links with these keywords
    links_to_remove = ["gsalr.com", "List", "estatesales", "salesestate", "companies", "real estate", "news"]
    if "estatesales" in url.lower():
    	links = get_relevant_estatesales_links(url)
    
    if "pennlive" in url.lower():
    	links = get_relevant_pennlive_links(url, keywords, links_to_remove)
    elif "auctionzip" in url.lower():
    	links = get_relevant_auctionzip_links(url)
    else:
    	links = get_relevant_links(url, keywords, links_to_remove)
    
    if not links:
        print(f"No relevant links found at {url}.")
        return addresses, descriptions, names, dates, url_list

    
    print(f"Found {len(links)} relevant links. Visiting each link...\n")
    print("--------------------------------------------------------------\n")

    for name, link in links:
        url_list.append(link)
        if "gsalr" in url:
            addresses, names, descriptions, dates = scrape_and_print_gsalr(link, addresses, names, descriptions, dates)
        if 'craigslist' in url:
            addresses, names, descriptions, dates = scrape_and_print_craigslist(link, addresses, names, descriptions, dates)
        if 'pennlive' in url:
        	addresses, names, descriptions, dates = scrape_and_print_pennlive(link, addresses, names, descriptions, dates)
        if 'auctionzip' in url:
        	addresses, names, descriptions, dates = scrape_and_print_auctionzip(link, addresses, names, descriptions, dates)



        print(f"Visited: {link}\n")
        print("*******************************************************\n")    
    
    
    return addresses, descriptions, names, dates,url_list

# Get next Saturday's date for Craigslist filtering
today = datetime.today()
days_until_saturday = (5 - today.weekday()) % 7
next_saturday = today + timedelta(days=days_until_saturday)
formatted_date = next_saturday.strftime('%Y-%m-%d')
year = (next_saturday.strftime('%Y'))
month = (next_saturday.strftime('%m'))
day = (next_saturday.strftime('%d'))

# List of URLs to scrape
main_urls = [
    "https://gsalr.com/garage-sales-harrisburg-pa.html?day=5",
    f"https://harrisburg.craigslist.org/search/gms?lat=40.2736&lon=-76.8847&sale_date={formatted_date}&search_distance=10",
#    f"https://www.auctionzip.com/cgi-bin/auctionlist.cgi?txtSearchZip=17055&txtSearchRadius=30&idxSearchCategory=0&gid=0&year={year}&month={month}&day={day}&txtSearchKeywords=&showlive=1",
    #"https://classifieds.pennlive.com/pennlive/category/garage-sale-estate-sale-auctions/garage-yard-estate-sales",
#    "https://garagesalefinder.com/yard-sales/mechanicsburg-pa/",
#    "https://garagesalefinder.com/yard-sales/camp-hill-pa/"
    ]

# Keywords to filter relevant links
keywords = ["Community", "Neighborhood", "Garage", "Yard", "Rummage", "Church", "Estate", "Development"]

addresses = []
descriptions = []
names = []
dates = []
url_list = []

# Run scraping on all URLs
for url in main_urls:
    addresses, descriptions, names, dates, url_list = main_scrape(url, keywords, addresses, descriptions, names, dates, url_list)
    
addresses, names, descriptions, dates, url_list = process_folder(addresses, names, descriptions, dates, url_list)

create_map(addresses, names, descriptions, dates, url_list)

export_data = {
    "generated_at": datetime.now().isoformat(),
    "event_count": len(export_events),
    "events": export_events
}

with open("yard_sales_latest.json", "w", encoding="utf-8") as f:
    json.dump(export_data, f, indent=4)

print(f"\nExported {len(export_events)} events to yard_sales_latest.json")