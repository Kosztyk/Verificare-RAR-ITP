"""RAR ITP Checker with Multiple Sensors"""
import logging
import aiohttp
from bs4 import BeautifulSoup
from datetime import datetime, timedelta, date
from homeassistant.components.sensor import SensorEntity
from homeassistant.helpers.update_coordinator import CoordinatorEntity, DataUpdateCoordinator, UpdateFailed
from homeassistant.util import slugify
from .const import DOMAIN, BASE_URL, DEFAULT_SCAN_INTERVAL, OCR_API_URL

_LOGGER = logging.getLogger(__name__)
SCAN_INTERVAL = timedelta(hours=DEFAULT_SCAN_INTERVAL)

# Month mapping for Romanian date parsing
MONTH_MAP = {
    "ian": "01", "feb": "02", "mar": "03", "apr": "04",
    "mai": "05", "iun": "06", "iul": "07", "aug": "08",
    "sept": "09", "oct": "10", "nov": "11", "dec": "12"
}

async def solve_captcha_with_ocrspace(image_bytes: bytes, api_key: str = None) -> str:
    """Solve CAPTCHA using OCR.Space API"""
    try:
        async with aiohttp.ClientSession() as session:
            form = aiohttp.FormData()
            form.add_field('file', image_bytes, filename='captcha.png')
            form.add_field('apikey', api_key or 'helloworld')
            form.add_field('language', 'eng')
            form.add_field('OCREngine', '2')
            async with session.post(OCR_API_URL, data=form) as resp:
                data = await resp.json()
                return data.get("ParsedResults", [{}])[0].get("ParsedText", "").strip()
    except Exception as e:
        _LOGGER.warning("OCR failed: %s", e)
        return ""

async def fetch_itp(vin: str, ocr_api_key: str = None) -> dict:
    """Fetch ITP data from RAR site."""
    async with aiohttp.ClientSession() as session:
        try:
            _LOGGER.info("Starting ITP check for VIN: %s", vin)
            # Initial GET to retrieve CAPTCHA image
            async with session.get(BASE_URL) as response:
                html = await response.text()
                soup = BeautifulSoup(html, "html.parser")

            captcha_img = soup.find("img", id="imgVerf")
            if not captcha_img or not captcha_img.get("src"):
                raise ValueError("CAPTCHA image not found")

            captcha_url = f"https://prog.rarom.ro/rarpol/{captcha_img['src']}"
            async with session.get(captcha_url, headers={"Referer": BASE_URL}) as cap_rsp:
                captcha_content = await cap_rsp.read()

            captcha_text = await solve_captcha_with_ocrspace(captcha_content, ocr_api_key)
            _LOGGER.debug("CAPTCHA solved: %s", captcha_text)

            form_data = {
                "serie_civ": "",
                "nr_id": vin,
                "cod_securitate": captcha_text,
                "trimite": "Caută"
            }
            async with session.post(BASE_URL, data=form_data) as result_response:
                result_text = await result_response.text()
                _LOGGER.debug("RAR ITP raw response for %s:\n%s", vin, result_text)

            # Parse only the result container if available
            result_soup = BeautifulSoup(result_text, "html.parser")
            result_div = result_soup.find('div', id='rezbgcolor')
            content_text = result_div.get_text(separator='\n', strip=True) if result_div else result_text
            lower = content_text.lower()

            # Default values
            status = "Not Found"
            expiration_date = "Unknown"

            # Check for captcha or no record errors
            if "codul de verificare a fost copiat incorect" in lower:
                raise UpdateFailed("Invalid CAPTCHA, please retry")

            if "nu a fost găsită nicio înregistrare" not in lower:
                status = "Valid"
                # New format parsing: 'valabilă până la d-mmm-yyyy'
                if "valabilă până la" in lower:
                    try:
                        fragment = lower.split("valabilă până la", 1)[1]
                        raw_date = fragment.split()[0].strip().strip('.')
                        day, month, year = raw_date.split("-")
                        expiration_date = f"{year}-{MONTH_MAP.get(month, '01')}-{day.zfill(2)}"
                        _LOGGER.debug("Parsed expiration_date: %s", expiration_date)
                    except Exception as e:
                        _LOGGER.warning("Failed to parse expiration date: %s", e)
                # Fallback old format parsing: 'Data expirării'</snip>
                elif "data expirării" in lower:
                    try:
                        node = result_soup.find(text=lambda t: "Data expirării" in t)
                        if node:
                            raw = node.find_next().get_text(strip=True)
                            day, month, year = raw.split('.')
                            expiration_date = f"{year}-{month}-{day}"
                            _LOGGER.debug("Parsed old-format expiration_date: %s", expiration_date)
                    except Exception as e:
                        _LOGGER.warning("Failed to parse old-format date: %s", e)

            return {
                "vin": vin,
                "status": status,
                "expiration_date": expiration_date,
                "last_checked": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
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
        super().__init__(coordinator)
        self.coordinator = coordinator
        vin = coordinator.data["vin"]
        self._attr_name = f"ITP Status {vin}"
        self._attr_unique_id = slugify(f"itp_status_{vin}")
        self._attr_icon = "mdi:car"

    @property
    def state(self):
        return self.coordinator.data.get("status", "unknown")

    @property
    def extra_state_attributes(self):
        return {
            "vin": self.coordinator.data.get("vin"),
            "last_checked": self.coordinator.data.get("last_checked"),
        }


class ITPExpirationDateSensor(CoordinatorEntity, SensorEntity):
    """ITP expiration date sensor."""
    _attr_device_class = "date"
    _attr_icon = "mdi:calendar-star"

    def __init__(self, coordinator):
        super().__init__(coordinator)
        self.coordinator = coordinator
        vin = coordinator.data["vin"]
        self._attr_name = f"ITP Expiration Date {vin}"
        self._attr_unique_id = slugify(f"itp_expiration_date_{vin}")

    @property
    def state(self):
        return self.coordinator.data.get("expiration_date", "Unknown")


class ITPLastCheckedSensor(CoordinatorEntity, SensorEntity):
    """Last checked timestamp sensor."""
    _attr_device_class = "timestamp"
    _attr_icon = "mdi:clock-outline"

    def __init__(self, coordinator):
        super().__init__(coordinator)
        self.coordinator = coordinator
        vin = coordinator.data["vin"]
        self._attr_name = f"ITP Last Checked {vin}"
        self._attr_unique_id = slugify(f"itp_last_checked_{vin}")

    @property
    def state(self):
        return self.coordinator.data.get("last_checked")


class ITPDaysLeftSensor(CoordinatorEntity, SensorEntity):
    """Days left until ITP expiration."""
    _attr_native_unit_of_measurement = "days"
    _attr_state_class = "measurement"
    _attr_icon = "mdi:calendar-clock"

    def __init__(self, coordinator):
        super().__init__(coordinator)
        self.coordinator = coordinator
        vin = coordinator.data["vin"]
        self._attr_name = f"ITP Days Left {vin}"
        self._attr_unique_id = slugify(f"itp_days_left_{vin}")

    @property
    def native_value(self):
        exp_date = self.coordinator.data.get("expiration_date")
        return calculate_days_until(exp_date)


async def async_setup_entry(hass, config_entry, async_add_entities):
    """Set up sensors from config entry."""
    vin = config_entry.data["vin"]
    ocr_key = config_entry.data.get("ocr_api_key", "")
    coordinator = DataUpdateCoordinator(
        hass,
        _LOGGER,
        name=f"{DOMAIN}_{vin}",
        update_method=lambda: fetch_itp(vin, ocr_key),
        update_interval=SCAN_INTERVAL,
    )
    hass.data.setdefault(DOMAIN, {})[vin] = {"coordinator": coordinator}
    await coordinator.async_config_entry_first_refresh()

    sensors = [
        ITPStatusSensor(coordinator),
        ITPExpirationDateSensor(coordinator),
        ITPLastCheckedSensor(coordinator),
        ITPDaysLeftSensor(coordinator),
    ]
    async_add_entities(sensors, True)
