from __future__ import annotations

import asyncio
import logging
import re
from time import monotonic

from benqprojector import BenQProjector
from homeassistant.components.select import SelectEntity, SelectEntityDescription
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant, callback
from homeassistant.helpers.entity import EntityCategory
from homeassistant.helpers.entity_platform import AddEntitiesCallback
from homeassistant.helpers.update_coordinator import CoordinatorEntity

from . import BenQProjectorCoordinator
from .const import DOMAIN

_LOGGER = logging.getLogger(__name__)


async def async_setup_entry(
    hass: HomeAssistant,
    config_entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up the BenQ Serial Projector select."""
    coordinator: BenQProjectorCoordinator = hass.data[DOMAIN][config_entry.entry_id]

    entity_descriptions = [
        SelectEntityDescription(
            key="audiosour",
            translation_key="audiosour",
            options=coordinator.projector.audio_sources or [],
        ),
        SelectEntityDescription(
            key="appmod",
            translation_key="appmod",
            options=coordinator.projector.picture_modes or [],
        ),
        SelectEntityDescription(
            key="ct",
            translation_key="ct",
            options=coordinator.projector.color_temperatures or [],
            entity_category=EntityCategory.CONFIG,
        ),
        SelectEntityDescription(
            key="asp",
            translation_key="asp",
            options=coordinator.projector.aspect_ratios or [],
        ),
        SelectEntityDescription(
            key="lampm",
            translation_key="lampm",
            options=coordinator.projector.lamp_modes or [],
            entity_category=EntityCategory.CONFIG,
        ),
        SelectEntityDescription(
            key="3d",
            translation_key="3d",
            options=coordinator.projector.threed_modes or [],
            entity_category=EntityCategory.CONFIG,
        ),
        # SelectEntityDescription(key="rr", None, translation_key="rr", entity_category=EntityCategory.CONFIG],
        SelectEntityDescription(
            key="pp",
            translation_key="pp",
            options=coordinator.projector.projector_positions or [],
            entity_category=EntityCategory.CONFIG,
        ),
        SelectEntityDescription(
            key="menuposition",
            translation_key="menuposition",
            options=coordinator.projector.menu_positions or [],
            entity_category=EntityCategory.CONFIG,
        ),
    ]

    entities = []

    for entity_description in entity_descriptions:
        if not entity_description.options:
            continue
        if coordinator.supports_command(
            entity_description.key, "read"
        ) and coordinator.supports_command(entity_description.key, "write"):
            entities.append(
                BenQProjectorSelect(
                    coordinator, entity_description, config_entry.entry_id
                )
            )

    async_add_entities(entities)


class BenQProjectorSelect(CoordinatorEntity, SelectEntity):
    _attr_has_entity_name = True
    _attr_available = False

    _attr_current_option = None

    def __init__(
        self,
        coordinator: BenQProjectorCoordinator,
        entity_description: SelectEntityDescription,
        config_entry_id: str,
    ) -> None:
        """Initialize the select."""
        super().__init__(coordinator, entity_description.key)

        self._attr_device_info = coordinator.device_info
        self._attr_unique_id = f"{config_entry_id}-{entity_description.key}"

        self._options_map = {
            re.sub("[^a-z0-9]", "_", value.lower()): value
            for value in (entity_description.options or [])
        }

        self.entity_description = entity_description
        self._last_probe_ts = 0.0
        self._probe_task: asyncio.Task | None = None

    def _normalize_option_key(self, value: str | None) -> str | None:
        """Map projector response value to HA select option key."""
        if value is None:
            return None

        if value in self._options_map:
            return value

        normalized = re.sub("[^a-z0-9]", "_", str(value).lower())
        if normalized in self._options_map:
            return normalized

        command_key = (self.entity_description.key or "").lower()
        aliases: dict[str, dict[str, str]] = {
            "appmod": {
                "cine": "cinema",
                "d_cinema": "d_cinema",
                "threed": "3d",
            },
            "lampm": {
                "lnor": "normal",
                "seco": "smarteco",
                "seco2": "smarteco2",
                "seco3": "smarteco3",
            },
            "asp": {
                "lbox": "letterbox",
            },
            "ct": {
                "native": "lamp_native",
            },
        }

        alias_key = aliases.get(command_key, {}).get(normalized)
        if alias_key and alias_key in self._options_map:
            return alias_key

        # Fuzzy fallback for vendor codes vs UI labels.
        for option_key in self._options_map:
            compact_option = option_key.replace("_", "")
            compact_value = normalized.replace("_", "")
            if compact_option in compact_value or compact_value in compact_option:
                return option_key

        # Last resort: accept the raw normalized value so entity doesn't stay unavailable.
        # Keep the original payload as display label when possible.
        self._options_map[normalized] = str(value)
        return normalized

        return None

    async def async_added_to_hass(self) -> None:
        await super().async_added_to_hass()

        if self.coordinator.data and (
            current_option := self.coordinator.data.get(self.entity_description.key)
        ):
            self._attr_current_option = self._normalize_option_key(current_option)
            self._attr_available = True
        else:
            _LOGGER.debug("%s is not available", self.entity_description.key)
            self._attr_available = False

        if self._attr_current_option is None:
            self._attr_available = False

        if (
            self._attr_current_option is None
            and self.coordinator.power_status
            in [
                BenQProjector.POWERSTATUS_POWERINGON,
                BenQProjector.POWERSTATUS_ON,
            ]
        ):
            await self._async_probe_current_option()

        self.async_write_ha_state()

    async def _async_probe_current_option(self) -> None:
        """Probe current option value when passive updates are noisy/missing."""
        now = monotonic()
        if (now - self._last_probe_ts) < 20:
            return
        self._last_probe_ts = now

        if self._probe_task and not self._probe_task.done():
            return

        async def _probe() -> None:
            try:
                response = await self.coordinator.async_send_command(
                    self.entity_description.key
                )
                normalized = self._normalize_option_key(response)
                if normalized is not None:
                    self._attr_current_option = normalized
                    self._attr_available = True
                    self.async_write_ha_state()
            except Exception:  # noqa: BLE001
                _LOGGER.debug(
                    "Probe failed for %s",
                    self.entity_description.key,
                )

        self._probe_task = self.hass.async_create_task(_probe())

    @property
    def available(self) -> bool:
        """Return if entity is available."""
        if not self._attr_available:
            return self._attr_available

        return self.coordinator.last_update_success

    @callback
    def _handle_coordinator_update(self) -> None:
        """Handle updated data from the coordinator."""
        if self.entity_description.key in self.coordinator.data:
            normalized = self._normalize_option_key(
                self.coordinator.data.get(self.entity_description.key)
            )
            if normalized is not None:
                self._attr_current_option = normalized
        has_value = self._attr_current_option is not None

        if self.coordinator.power_status == BenQProjector.POWERSTATUS_UNKNOWN:
            self._attr_available = False
        elif self.coordinator.power_status in [
            BenQProjector.POWERSTATUS_POWERINGON,
            BenQProjector.POWERSTATUS_ON,
        ]:
            self._attr_available = has_value
            if not has_value:
                self.hass.async_create_task(self._async_probe_current_option())
        else:
            self._attr_available = False

        self.async_write_ha_state()

    @property
    def options(self) -> list[str]:
        return list(self._options_map.keys())

    async def async_select_option(self, option: str) -> None:
        """Change the selected option."""
        option = self._options_map[option]

        response = await self.coordinator.async_send_command(
            self.entity_description.key, option
        )
        if response == option:
            self._attr_current_option = option
            self.async_write_ha_state()
        else:
            _LOGGER.error("Failed to set %s to %s", self.name, option)
