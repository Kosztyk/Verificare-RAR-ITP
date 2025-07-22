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
    """Solve CAPTCHA using OCR.Space API with improved error handling"""
    try:
        async with aiohttp.ClientSession() as session:
            form = aiohttp.FormData()
            form.add_field("file", image_bytes, filename="captcha.png")
            form.add_field("apikey", api_key or "helloworld")
            form.add_field("language", "eng")
            form.add_field("OCREngine", "2")

            try:
                async with session.post(OCR_API_URL, data=form) as resp:
                    # Handle non-JSON responses
                    if "application/json" not in resp.content_type:
                        text = await resp.text()
                        _LOGGER.warning(
                            "OCR API returned non-JSON: %s %s", resp.status, text[:100]
                        )
                        raise OCRAPIError("Non-JSON response from OCR API")

                    data = await resp.json()
                    if resp.status != 200:
                        error_msg = data.get("ErrorMessage", "Unknown error")
                        _LOGGER.warning("OCR API error: %s", error_msg)
                        raise OCRAPIError(f"OCR API error: {error_msg}")

                    result = data.get("ParsedResults", [{}])[0].get("ParsedText", "").strip()
                    # Validate it's a 4-6 digit code
                    if result and re.match(r"^\d{4,6}$", result):
                        return result
                    raise OCRAPIError("Invalid CAPTCHA format")

            except (aiohttp.ClientError, asyncio.TimeoutError) as e:
                _LOGGER.warning("OCR API request failed: %s", str(e))
                raise OCRAPIError("API request failed") from e

    except Exception as e:
        _LOGGER.warning("OCR processing failed: %s", str(e))
        raise OCRAPIError("OCR processing failed") from e


async def fetch_itp(vin: str, ocr_api_key: str = None) -> dict:
    """Fetch ITP data from RAR site with improved error handling."""
    timeout = aiohttp.ClientTimeout(total=30)  # 30 second timeout

    async with aiohttp.ClientSession(timeout=timeout) as session:
        try:
            _LOGGER.info("Starting ITP check for VIN: %s", vin)

            # Retry loop for CAPTCHA
            for attempt in range(3):
                try:
                    # Initial GET with timeout
                    try:
                        async with session.get(BASE_URL) as response:
                            html = await response.text()
                    except asyncio.TimeoutError:
                        raise UpdateFailed("Timeout connecting to RAR website")

                    soup = BeautifulSoup(html, "html.parser")
                    captcha_img = soup.find("img", id="imgVerf")
                    if not captcha_img or not captcha_img.get("src"):
                        raise ValueError("CAPTCHA image not found")

                    # Download CAPTCHA with timeout
                    try:
                        captcha_url = f"https://prog.rarom.ro/rarpol/{captcha_img['src']}"
                        async with session.get(
                            captcha_url, headers={"Referer": BASE_URL}
                        ) as cap_rsp:
                            captcha_content = await cap_rsp.read()
                    except asyncio.TimeoutError:
                        raise UpdateFailed("Timeout downloading CAPTCHA image")

                    try:
                        captcha_text = await solve_captcha_with_ocrspace(
                            captcha_content, ocr_api_key
                        )
                    except OCRAPIError as e:
                        _LOGGER.warning("CAPTCHA solving failed: %s", str(e))
                        if attempt == 2:  # Last attempt
                            raise UpdateFailed("Failed to solve CAPTCHA after 3 attempts")
                        continue

                    clean_captcha = re.sub(r"\D", "", captcha_text)
                    form_data = {
                        "serie_civ": "",
                        "nr_id": vin.upper(),
                        "verif_cod": clean_captcha,
                        "trimite": "Caută",
                        "from_url": "",
                        "id": "",
                    }
                    headers = {
                        "Referer": BASE_URL,
                        "Origin": "https://prog.rarom.ro",
                        "User-Agent": "Mozilla/5.0 (HA RAR ITP Checker)",
                    }

                    _LOGGER.debug(
                        "Form data being submitted: %s",
                        {k: v for k, v in form_data.items() if k != "verif_cod"},
                    )

                    # Submit form with timeout
                    try:
                        async with session.post(
                            BASE_URL, data=form_data, headers=headers
                        ) as result_response:
                            result_text = await result_response.text()
                            _LOGGER.debug(
                                "Response status: %s, content-type: %s",
                                result_response.status,
                                result_response.content_type,
                            )
                    except asyncio.TimeoutError:
                        raise UpdateFailed("Timeout submitting form to RAR website")

                    # Check if CAPTCHA was accepted
                    if "codul de verificare a fost copiat incorect" in result_text.lower():
                        if attempt == 2:  # Last attempt
                            raise UpdateFailed("CAPTCHA validation failed after 3 attempts")
                        continue  # Try again

                    break  # Success - proceed with parsing

                except UpdateFailed as uf:
                    if attempt == 2:  # Last attempt
                        raise
                    _LOGGER.debug("Attempt %d failed, retrying: %s", attempt + 1, uf)
                    await asyncio.sleep(2)  # Add delay between retries
                    continue

            # Parse only the result container if available
            result_soup = BeautifulSoup(result_text, "html.parser")
            result_div = result_soup.find("div", id="rezbgcolor")
            content_text = (
                result_div.get_text(separator="\n", strip=True)
                if result_div
                else result_text
            )
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
                        _LOGGER.debug("Parsed expiration_date: %s", expiration_date)
                    except Exception as e:
                        _LOGGER.warning("Failed to parse expiration date: %s", e)
                # Fallback old format parsing: 'Data expirării'
                elif "data expirării" in lower:
                    try:
                        node = result_soup.find(text=lambda t: "Data expirării" in t)
                        if node:
                            raw = node.find_next().get_text(strip=True)
                            day, month, year = raw.split(".")
                            expiration_date = f"{year}-{month}-{day}"
                            _LOGGER.debug("Parsed old-format expiration_date: %s", expiration_date)
                    except Exception as e:
                        _LOGGER.warning("Failed to parse old-format date: %s", e)

            return {
                "vin": vin,
                "status": status,
                "expiration_date": expiration_date,
                "last_checked": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

        except UpdateFailed as uf:
            _LOGGER.error("ITP check failed for %s: %s", vin, uf)
            raise
        except Exception as err:
            _LOGGER.error("ITP check error for %s: %s", vin, err)
            raise UpdateFailed(f"ITP check failed: {err}")


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
                await asyncio.sleep(2)  # Add delay between retries
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