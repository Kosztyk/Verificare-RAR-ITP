"""RAR ITP Checker with Multiple Sensors"""
import asyncio
import logging
import re
import aiohttp
from bs4 import BeautifulSoup
from datetime import datetime, timedelta, date
from homeassistant.components.sensor import SensorEntity
from homeassistant.helpers.update_coordinator import (
    CoordinatorEntity,
    DataUpdateCoordinator,
    UpdateFailed,
)
from homeassistant.exceptions import ConfigEntryNotReady
from homeassistant.util import slugify
from .const import DOMAIN, BASE_URL, DEFAULT_SCAN_INTERVAL, OCR_API_URL

_LOGGER = logging.getLogger(__name__)
SCAN_INTERVAL = timedelta(hours=DEFAULT_SCAN_INTERVAL)

# Month mapping for Romanian date parsing
MONTH_MAP = {
    "ian": "01",
    "feb": "02",
    "mar": "03",
    "apr": "04",
    "mai": "05",
    "iun": "06",
    "iul": "07",
    "aug": "08",
    "sept": "09",
    "oct": "10",
    "nov": "11",
    "dec": "12",
}

class OCRAPIError(Exception):
    """Custom exception for OCR API errors"""

async def solve_captcha_with_ocrspace(image_bytes: bytes, api_key: str = None) -> str:
    """Solve CAPTCHA using OCR.Space API with improved timeout handling"""
    try:
        async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=15)) as session:
            form = aiohttp.FormData()
            form.add_field("file", image_bytes, filename="captcha.png")
            form.add_field("apikey", api_key or "helloworld")  # fallback to free tier
            form.add_field("language", "eng")
            form.add_field("OCREngine", "2")  # Best for CAPTCHAs
            form.add_field("isOverlayRequired", "false")

            try:
                async with session.post(OCR_API_URL, data=form) as resp:
                    if resp.status != 200:
                        error_msg = f"OCR API returned status {resp.status}"
                        _LOGGER.warning(error_msg)
                        raise OCRAPIError(error_msg)

                    data = await resp.json()
                    if not data.get("ParsedResults"):
                        error_msg = data.get("ErrorMessage", "No parsed results")
                        if isinstance(error_msg, list):
                            error_msg = ", ".join(error_msg)
                        _LOGGER.warning("OCR API error: %s", error_msg)
                        raise OCRAPIError(f"OCR failed: {error_msg}")

                    result = data["ParsedResults"][0].get("ParsedText", "").strip()
                    if not re.match(r"^\d{4,6}$", result):  # Validate CAPTCHA format
                        raise OCRAPIError(f"Invalid CAPTCHA format: {result}")
                    
                    return result

            except asyncio.TimeoutError:
                _LOGGER.warning("OCR API timeout, retrying with longer timeout")
                # Retry with longer timeout
                async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=30)) as retry_session:
                    async with retry_session.post(OCR_API_URL, data=form) as resp:
                        data = await resp.json()
                        if not data.get("ParsedResults"):
                            raise OCRAPIError("OCR failed after retry")
                        return data["ParsedResults"][0].get("ParsedText", "").strip()

    except Exception as e:
        _LOGGER.warning("OCR processing failed: %s", str(e))
        raise OCRAPIError("OCR processing failed") from e

async def fetch_itp(vin: str, ocr_api_key: str = None) -> dict:
    """Fetch ITP data from RAR site with robust CAPTCHA handling."""
    timeout = aiohttp.ClientTimeout(total=30)
    headers = {
        "User-Agent": "Mozilla/5.0 (HA RAR ITP Checker)",
        "Referer": BASE_URL,
        "Origin": "https://prog.rarom.ro"
    }

    async with aiohttp.ClientSession(timeout=timeout, headers=headers) as session:
        try:
            _LOGGER.info("Starting ITP check for VIN: %s", vin)

            # Initial page load
            async with session.get(BASE_URL) as response:
                if response.status != 200:
                    raise UpdateFailed(f"Initial request failed: HTTP {response.status}")
                html = await response.text()

            soup = BeautifulSoup(html, "html.parser")
            
            # CAPTCHA handling with retries
            for attempt in range(3):
                try:
                    # Locate CAPTCHA image
                    captcha_img = soup.find("img", id="imgVerf")
                    if not captcha_img or not captcha_img.get("src"):
                        _LOGGER.debug("CAPTCHA HTML: %s", str(captcha_img))
                        raise UpdateFailed("CAPTCHA image not found in page")

                    # Build CAPTCHA URL
                    captcha_src = captcha_img['src']
                    if captcha_src.startswith("http"):
                        captcha_url = captcha_src
                    else:
                        captcha_url = f"https://prog.rarom.ro/rarpol/{captcha_src.lstrip('/')}"
                    
                    _LOGGER.debug("Downloading CAPTCHA from: %s", captcha_url)
                    
                    # Download CAPTCHA image
                    async with session.get(captcha_url) as cap_resp:
                        if cap_resp.status != 200:
                            raise UpdateFailed(f"CAPTCHA download failed: HTTP {cap_resp.status}")
                        captcha_content = await cap_resp.read()

                    # Solve CAPTCHA with retry logic
                    try:
                        captcha_text = await solve_captcha_with_ocrspace(captcha_content, ocr_api_key)
                    except OCRAPIError as e:
                        if attempt == 2:  # Last attempt
                            raise
                        await asyncio.sleep(2)
                        continue

                    clean_captcha = re.sub(r"\D", "", captcha_text)  # Keep only digits
                    _LOGGER.debug("CAPTCHA solved: %s", clean_captcha)

                    # Prepare form data
                    form_data = {
                        "serie_civ": "",
                        "nr_id": vin.upper(),
                        "antirobot": clean_captcha,
                        "trimite": "Caută",
                        "from_url": "",
                        "id": "",
                    }

                    # Submit form
                    async with session.post(BASE_URL, data=form_data) as result_response:
                        result_text = await result_response.text()
                        
                        if "codul de verificare a fost copiat incorect" in result_text.lower():
                            raise UpdateFailed("CAPTCHA validation failed")
                        
                        # Success - proceed to parse results
                        break

                except (UpdateFailed, OCRAPIError) as e:
                    if attempt == 2:  # Last attempt
                        raise UpdateFailed(f"Failed after 3 attempts: {str(e)}")
                    _LOGGER.debug("Attempt %d failed, retrying: %s", attempt + 1, e)
                    await asyncio.sleep(2)
                    continue

            # Parse results
            result_soup = BeautifulSoup(result_text, "html.parser")
            result_div = result_soup.find("div", id="rezbgcolor")
            content_text = result_div.get_text(separator="\n", strip=True) if result_div else result_text
            lower = content_text.lower()

            # Default values
            status = "Not Found"
            expiration_date = "Unknown"

            if "nu a fost găsită nicio înregistrare" not in lower:
                status = "Valid"
                # New format parsing: 'valabilă până la d-mmm-yyyy'
                if "valabilă până la" in lower:
                    try:
                        fragment = lower.split("valabilă până la", 1)[1]
                        raw_date = fragment.split()[0].strip().strip(".")
                        day, month, year = raw_date.split("-")
                        expiration_date = f"{year}-{MONTH_MAP.get(month, '01')}-{day.zfill(2)}"
                    except Exception as e:
                        _LOGGER.warning("Failed to parse expiration date: %s", e)
                # Fallback old format parsing
                elif "data expirării" in lower:
                    try:
                        node = result_soup.find(text=lambda t: "Data expirării" in t)
                        if node:
                            raw = node.find_next().get_text(strip=True)
                            day, month, year = raw.split(".")
                            expiration_date = f"{year}-{month}-{day}"
                    except Exception as e:
                        _LOGGER.warning("Failed to parse old-format date: %s", e)

            return {
                "vin": vin,
                "status": status,
                "expiration_date": expiration_date,
                "last_checked": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

        except Exception as err:
            _LOGGER.error("ITP check failed for %s: %s", vin, err, exc_info=True)
            raise UpdateFailed(f"ITP check failed: {err}") from err

def calculate_days_until(expiration_date: str) -> int | None:
    """Calculate days until expiration."""
    if not expiration_date or expiration_date == "Unknown":
        return None
    try:
        exp = datetime.strptime(expiration_date, "%Y-%m-%d").date()
        return (exp - date.today()).days
    except ValueError:
        return None

class ITPStatusSensor(CoordinatorEntity, SensorEntity):
    """ITP status sensor."""

    def __init__(self, coordinator):
        """Initialize the sensor."""
        super().__init__(coordinator)
        self.coordinator = coordinator
        vin = coordinator.data["vin"]
        self._attr_name = f"ITP Status {vin}"
        self._attr_unique_id = slugify(f"itp_status_{vin}")
        self._attr_icon = "mdi:car"

    @property
    def state(self):
        """Return the state of the sensor."""
        return self.coordinator.data.get("status", "unknown")

    @property
    def extra_state_attributes(self):
        """Return additional state attributes."""
        return {
            "vin": self.coordinator.data.get("vin"),
            "last_checked": self.coordinator.data.get("last_checked"),
        }

class ITPExpirationDateSensor(CoordinatorEntity, SensorEntity):
    """ITP expiration date sensor."""

    _attr_device_class = "date"
    _attr_icon = "mdi:calendar-star"

    def __init__(self, coordinator):
        """Initialize the sensor."""
        super().__init__(coordinator)
        self.coordinator = coordinator
        vin = coordinator.data["vin"]
        self._attr_name = f"ITP Expiration Date {vin}"
        self._attr_unique_id = slugify(f"itp_expiration_date_{vin}")

    @property
    def state(self):
        """Return the state of the sensor."""
        return self.coordinator.data.get("expiration_date", "Unknown")

class ITPLastCheckedSensor(CoordinatorEntity, SensorEntity):
    """Last checked timestamp sensor."""

    _attr_device_class = "timestamp"
    _attr_icon = "mdi:clock-outline"

    def __init__(self, coordinator):
        """Initialize the sensor."""
        super().__init__(coordinator)
        self.coordinator = coordinator
        vin = coordinator.data["vin"]
        self._attr_name = f"ITP Last Checked {vin}"
        self._attr_unique_id = slugify(f"itp_last_checked_{vin}")

    @property
    def state(self):
        """Return the state of the sensor."""
        return self.coordinator.data.get("last_checked")

class ITPDaysLeftSensor(CoordinatorEntity, SensorEntity):
    """Days left until ITP expiration."""

    _attr_native_unit_of_measurement = "days"
    _attr_state_class = "measurement"
    _attr_icon = "mdi:calendar-clock"

    def __init__(self, coordinator):
        """Initialize the sensor."""
        super().__init__(coordinator)
        self.coordinator = coordinator
        vin = coordinator.data["vin"]
        self._attr_name = f"ITP Days Left {vin}"
        self._attr_unique_id = slugify(f"itp_days_left_{vin}")

    @property
    def native_value(self):
        """Return the native value of the sensor."""
        exp_date = self.coordinator.data.get("expiration_date")
        return calculate_days_until(exp_date)

async def async_setup_entry(hass, config_entry, async_add_entities):
    """Set up sensors from config entry with improved error handling."""
    vin = config_entry.data["vin"]
    ocr_key = config_entry.data.get("ocr_api_key", "")

    async def async_update_data():
        """Wrap the fetch with retry logic."""
        for attempt in range(3):
            try:
                return await fetch_itp(vin, ocr_key)
            except UpdateFailed as e:
                if attempt == 2:  # Last attempt
                    raise
                _LOGGER.debug("Attempt %d failed, retrying: %s", attempt + 1, e)
                await asyncio.sleep(2)
                continue

    coordinator = DataUpdateCoordinator(
        hass,
        _LOGGER,
        name=f"{DOMAIN}_{vin}",
        update_method=async_update_data,
        update_interval=SCAN_INTERVAL,
    )

    try:
        await coordinator.async_config_entry_first_refresh()
    except Exception as ex:
        _LOGGER.error("Failed to setup RAR ITP Checker: %s", str(ex))
        raise ConfigEntryNotReady from ex

    hass.data.setdefault(DOMAIN, {})[vin] = {"coordinator": coordinator}

    sensors = [
        ITPStatusSensor(coordinator),
        ITPExpirationDateSensor(coordinator),
        ITPLastCheckedSensor(coordinator),
        ITPDaysLeftSensor(coordinator),
    ]
    async_add_entities(sensors, True)