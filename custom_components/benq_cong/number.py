from __future__ import annotations

import asyncio
import logging
from time import monotonic

from benqprojector import BenQProjector
from homeassistant.components.number import NumberEntity, NumberEntityDescription
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
    """Set up the BenQ Serial Projector number."""
    coordinator: BenQProjectorCoordinator = hass.data[DOMAIN][config_entry.entry_id]

    entity_descriptions = [
        NumberEntityDescription(key="con", translation_key="con", native_max_value=100),
        NumberEntityDescription(key="bri", translation_key="bri", native_max_value=100),
        NumberEntityDescription(
            key="color", translation_key="color", native_max_value=20
        ),
        NumberEntityDescription(
            key="sharp", translation_key="sharp", native_max_value=20
        ),
        NumberEntityDescription(
            key="micvol", translation_key="micvol", native_max_value=20
        ),
        NumberEntityDescription(
            key="keyst",
            translation_key="keyst",
            native_max_value=20,
            entity_category=EntityCategory.CONFIG,
            entity_registry_enabled_default=False,
        ),
        NumberEntityDescription(
            key="hkeystone",
            translation_key="hkeystone",
            native_max_value=20,
            entity_category=EntityCategory.CONFIG,
            entity_registry_enabled_default=False,
        ),
        NumberEntityDescription(
            key="vkeystone",
            translation_key="vkeystone",
            native_max_value=20,
            entity_category=EntityCategory.CONFIG,
            entity_registry_enabled_default=False,
        ),
        NumberEntityDescription(
            key="rgain",
            translation_key="rgain",
            native_max_value=200,
            entity_category=EntityCategory.CONFIG,
            entity_registry_enabled_default=False,
        ),
        NumberEntityDescription(
            key="ggain",
            translation_key="ggain",
            native_max_value=200,
            entity_category=EntityCategory.CONFIG,
            entity_registry_enabled_default=False,
        ),
        NumberEntityDescription(
            key="bgain",
            translation_key="bgain",
            native_max_value=200,
            entity_category=EntityCategory.CONFIG,
            entity_registry_enabled_default=False,
        ),
        NumberEntityDescription(
            key="roffset",
            translation_key="roffset",
            native_max_value=511,
            entity_category=EntityCategory.CONFIG,
            entity_registry_enabled_default=False,
        ),
        NumberEntityDescription(
            key="goffset",
            translation_key="goffset",
            native_max_value=511,
            entity_category=EntityCategory.CONFIG,
            entity_registry_enabled_default=False,
        ),
        NumberEntityDescription(
            key="boffset",
            translation_key="boffset",
            native_max_value=511,
            entity_category=EntityCategory.CONFIG,
            entity_registry_enabled_default=False,
        ),
        # NumberEntityDescription(key="gamma", translation_key="gamma", native_min_value=1.6, native_max_value=2.8, native_step=0.1, entity_category=EntityCategory.CONFIG, entity_registry_enabled_default=False,),
        NumberEntityDescription(
            key="hdrbri",
            translation_key="hdrbri",
            native_min_value=-2,
            native_max_value=2,
            entity_category=EntityCategory.CONFIG,
            entity_registry_enabled_default=False,
        ),
    ]

    entities = []

    for entity_description in entity_descriptions:
        if coordinator.supports_command(
            entity_description.key, "read"
        ) and coordinator.supports_command(entity_description.key, "write"):
            entities.append(
                BenQProjectorNumber(
                    coordinator, entity_description, config_entry.entry_id
                )
            )

    async_add_entities(entities)


class BenQProjectorNumber(CoordinatorEntity, NumberEntity):
    _attr_has_entity_name = True
    _attr_available = False
    _attr_native_min_value = 0
    _attr_native_step = 1
    _attr_native_value = None

    def __init__(
        self,
        coordinator: BenQProjectorCoordinator,
        entity_description: NumberEntityDescription,
        config_entry_id: str,
    ) -> None:
        """Initialize the number."""
        super().__init__(coordinator, entity_description.key)

        self._attr_device_info = coordinator.device_info
        self._attr_unique_id = f"{config_entry_id}-{entity_description.key}"

        self.entity_description = entity_description
        self._last_probe_ts = 0.0
        self._probe_task: asyncio.Task | None = None

    async def _async_probe_value(self) -> None:
        """Probe current numeric value when passive updates are missing."""
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
                if response in (None, "", "?", "unknown", "unavailable"):
                    return

                self._attr_native_value = float(response)
                self._attr_available = True
                self.async_write_ha_state()
            except (ValueError, TypeError):
                _LOGGER.debug(
                    "Ignoring non-numeric probe %s value: %s",
                    self.entity_description.key,
                    response if "response" in locals() else None,
                )
            except Exception:  # noqa: BLE001
                _LOGGER.debug("Probe failed for %s", self.entity_description.key)

        self._probe_task = self.hass.async_create_task(_probe())

    async def async_added_to_hass(self) -> None:
        await super().async_added_to_hass()

        if self.coordinator.data and (
            native_value := self.coordinator.data.get(self.entity_description.key)
        ):
            try:
                self._attr_native_value = float(native_value)
                self._attr_available = True
                self.async_write_ha_state()
            except ValueError as ex:
                _LOGGER.error(
                    "ValueError for %s = %s, %s",
                    self.entity_description.key,
                    self.coordinator.data.get(self.entity_description.key),
                    ex,
                )
            except TypeError as ex:
                _LOGGER.error("TypeError for %s, %s", self.entity_description.key, ex)
        else:
            _LOGGER.debug("%s is not available", self.entity_description.key)

        if (
            self._attr_native_value is None
            and self.coordinator.power_status
            in [
                BenQProjector.POWERSTATUS_POWERINGON,
                BenQProjector.POWERSTATUS_ON,
            ]
        ):
            await self._async_probe_value()

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
            try:
                self._attr_native_value = float(
                    self.coordinator.data.get(self.entity_description.key)
                )
            except ValueError:
                _LOGGER.debug(
                    "Ignoring non-numeric %s value: %s",
                    self.entity_description.key,
                    self.coordinator.data.get(self.entity_description.key),
                )
            except TypeError:
                _LOGGER.debug(
                    "Ignoring invalid %s value type: %s",
                    self.entity_description.key,
                    type(self.coordinator.data.get(self.entity_description.key)).__name__,
                )

        has_value = self._attr_native_value is not None

        if self.coordinator.power_status == BenQProjector.POWERSTATUS_UNKNOWN:
            self._attr_available = False
        elif self.coordinator.power_status in [
            BenQProjector.POWERSTATUS_POWERINGON,
            BenQProjector.POWERSTATUS_ON,
        ]:
            self._attr_available = has_value
            if not has_value:
                self.hass.async_create_task(self._async_probe_value())
        else:
            self._attr_available = False

        self.async_write_ha_state()

    async def async_set_native_value(self, value: float) -> None:
        if self.coordinator.power_status == BenQProjector.POWERSTATUS_ON:
            if self._attr_native_value == value:
                return

            while self._attr_native_value < value:
                if (
                    await self.coordinator.async_send_command(
                        self.entity_description.key, "+"
                    )
                    == "+"
                ):
                    self._attr_native_value += self._attr_native_step
                else:
                    break

            while self._attr_native_value > value:
                if (
                    await self.coordinator.async_send_command(
                        self.entity_description.key, "-"
                    )
                    == "-"
                ):
                    self._attr_native_value -= self._attr_native_step
                else:
                    break

            self.async_write_ha_state()
        else:
            self._attr_available = False

        self.async_write_ha_state()
