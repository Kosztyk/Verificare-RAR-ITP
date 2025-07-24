from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant

from .const import DOMAIN

async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry):
    hass.data.setdefault(DOMAIN, {})
    
    await hass.config_entries.async_forward_entry_setups(entry, ["sensor"])

    async def handle_check_now(call):
        vin = entry.data["vin"]
        if vin in hass.data[DOMAIN]:
            coordinator = hass.data[DOMAIN][vin]["coordinator"]
            await coordinator.async_request_refresh()

    hass.services.async_register(DOMAIN, "check_now", handle_check_now)

    return True

async def async_unload_entry(hass: HomeAssistant, entry: ConfigEntry):
    """Unload a config entry."""
    unload_ok = await hass.config_entries.async_unload_platforms(entry, ["sensor"])
    if unload_ok:
        hass.data[DOMAIN].pop(entry.data["vin"])
        if not hass.data[DOMAIN]:
            hass.services.async_remove(DOMAIN, "check_now")
    return unload_ok