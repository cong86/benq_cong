from __future__ import annotations

from dataclasses import dataclass

from homeassistant.components.button import ButtonEntity, ButtonEntityDescription
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.entity_platform import AddEntitiesCallback
from homeassistant.helpers.update_coordinator import CoordinatorEntity

from . import BenQProjectorCoordinator
from .const import DOMAIN


@dataclass(frozen=True)
class BenQProjectorPowerButtonDescription(ButtonEntityDescription):
    """Describe a BenQ projector power button entity."""

    action: str = ""


POWER_BUTTONS: tuple[BenQProjectorPowerButtonDescription, ...] = (
    BenQProjectorPowerButtonDescription(
        key="power_on",
        translation_key="power_on",
        action="turn_on",
    ),
    BenQProjectorPowerButtonDescription(
        key="power_off",
        translation_key="power_off",
        action="turn_off",
    ),
)


async def async_setup_entry(
    hass: HomeAssistant,
    config_entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up BenQ projector power buttons."""
    coordinator: BenQProjectorCoordinator = hass.data[DOMAIN][config_entry.entry_id]

    entities = [
        BenQProjectorPowerButton(coordinator, description, config_entry.entry_id)
        for description in POWER_BUTTONS
    ]
    async_add_entities(entities)


class BenQProjectorPowerButton(CoordinatorEntity, ButtonEntity):
    """BenQ projector power action button."""

    _attr_has_entity_name = True

    def __init__(
        self,
        coordinator: BenQProjectorCoordinator,
        entity_description: BenQProjectorPowerButtonDescription,
        config_entry_id: str,
    ) -> None:
        super().__init__(coordinator)
        self.entity_description = entity_description
        self._attr_device_info = coordinator.device_info
        self._attr_unique_id = f"{config_entry_id}-{entity_description.key}"

    async def async_press(self) -> None:
        """Handle button press."""
        if self.entity_description.action == "turn_on":
            await self.coordinator.async_turn_on()
            return

        await self.coordinator.async_turn_off()
