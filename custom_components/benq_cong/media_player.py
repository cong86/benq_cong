"""The BenQ Projector integration."""

from __future__ import annotations

import logging

from benqprojector import BenQProjector
from homeassistant.components.media_player import (
    MediaPlayerDeviceClass,
    MediaPlayerEntity,
    MediaPlayerEntityFeature,
    MediaPlayerState,
)
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant, callback
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
    """Set up the BenQ Projector media player."""
    coordinator: BenQProjectorCoordinator = hass.data[DOMAIN][config_entry.entry_id]

    async_add_entities([BenQProjectorMediaPlayer(coordinator, config_entry.entry_id)])


class BenQProjectorMediaPlayer(CoordinatorEntity, MediaPlayerEntity):
    _attr_has_entity_name = True
    _attr_name = None
    _attr_device_class = MediaPlayerDeviceClass.TV
    _attr_translation_key = "projector"
    _attr_supported_features = (
        MediaPlayerEntityFeature.VOLUME_MUTE
        | MediaPlayerEntityFeature.VOLUME_SET
        | MediaPlayerEntityFeature.VOLUME_STEP
        | MediaPlayerEntityFeature.TURN_ON
        | MediaPlayerEntityFeature.TURN_OFF
        | MediaPlayerEntityFeature.SELECT_SOURCE
    )

    _attr_available = False
    _attr_state = None

    _attr_source_list = None
    _attr_source = None

    _attr_is_volume_muted = None
    _attr_volume_level = None

    def __init__(
        self, coordinator: BenQProjectorCoordinator, config_entry_id: str
    ) -> None:
        """Initialize the media player."""
        super().__init__(coordinator)

        self._attr_device_info = coordinator.device_info
        self._attr_unique_id = f"{config_entry_id}-projector"

    def _get_source_translation_key(self, source: str):
        """
        Projectors can have 1 or multiple sources for HDMI, RGB and YPBR. In case multiple sources
        of the same kind are present the source translation should include a sequence number, if
        only one source of a kind is present no sequence number is needed in the translation.
        """
        source_translation_key = source
        if (
            source in ("hdmi", "rgb", "ypbr")
            and len([s for s in self._attr_source_list if s.startswith(source)]) > 1
        ):
            # More than 1 source of this kind present, add "1" to the source to use the translation
            # with sequence number.
            source_translation_key = source + "1"
        return source_translation_key

    @staticmethod
    def _normalize_volume_level(raw_value) -> float | None:
        """Convert projector volume payload to HA 0..1 float."""
        if raw_value in (None, "", "?", "unknown", "unavailable"):
            return None

        try:
            volume = float(raw_value)
        except (TypeError, ValueError):
            return None

        # BenQ volume scale is typically 0..20.
        return max(0.0, min(volume / 20.0, 1.0))

    @staticmethod
    def _normalize_mute(raw_value) -> bool | None:
        """Convert projector mute payload to boolean."""
        if raw_value is None:
            return None

        if isinstance(raw_value, bool):
            return raw_value

        text = str(raw_value).strip().lower()
        if text in ("on", "1", "true", "yes"):
            return True
        if text in ("off", "0", "false", "no"):
            return False

        return None

    async def async_added_to_hass(self) -> None:
        await super().async_added_to_hass()

        self._attr_source_list = self.coordinator.video_sources or []
        source_list = [
            self._get_source_translation_key(source)
            for source in (self.coordinator.video_sources or [])
        ]
        self._attr_source_list = source_list

        if self.coordinator.power_status == BenQProjector.POWERSTATUS_UNKNOWN:
            _LOGGER.debug("Projector is not available")
            self._attr_available = False
        elif self.coordinator.power_status in [
            BenQProjector.POWERSTATUS_POWERINGON,
            BenQProjector.POWERSTATUS_ON,
        ]:
            self._attr_state = MediaPlayerState.ON

            volume_level = self._normalize_volume_level(self.coordinator.volume)
            if volume_level is not None:
                self._attr_volume_level = volume_level

            muted = self._normalize_mute(self.coordinator.muted)
            if muted is not None:
                self._attr_is_volume_muted = muted

            if self.coordinator.video_source is not None:
                self._attr_source = self._get_source_translation_key(
                    self.coordinator.video_source
                )

            self._attr_available = True
        elif self.coordinator.power_status == BenQProjector.POWERSTATUS_POWERINGOFF:
            self._attr_state = MediaPlayerState.OFF
            self._attr_available = False
        elif self.coordinator.power_status == BenQProjector.POWERSTATUS_OFF:
            self._attr_state = MediaPlayerState.OFF
            self._attr_available = True

        self.async_write_ha_state()

    @property
    def available(self) -> bool:
        """Return if entity is available."""
        if not self._attr_available:
            return self._attr_available

        return self.coordinator.last_update_success

    @callback
    def _handle_coordinator_update(self) -> None:
        """Handle updated data from the coordinator."""
        if self.coordinator.power_status == BenQProjector.POWERSTATUS_UNKNOWN:
            self._attr_available = False
        elif self.coordinator.power_status in [
            BenQProjector.POWERSTATUS_POWERINGON,
            BenQProjector.POWERSTATUS_ON,
        ]:
            self._attr_state = MediaPlayerState.ON
            self._attr_available = True
        elif self.coordinator.power_status == BenQProjector.POWERSTATUS_POWERINGOFF:
            self._attr_state = MediaPlayerState.OFF
            self._attr_available = False
        elif self.coordinator.power_status == BenQProjector.POWERSTATUS_OFF:
            self._attr_state = MediaPlayerState.OFF
            self._attr_available = True

        if "vol" in self.coordinator.data:
            volume_level = self._normalize_volume_level(self.coordinator.data.get("vol"))
            if volume_level is not None:
                self._attr_volume_level = volume_level

        if "mute" in self.coordinator.data:
            muted = self._normalize_mute(self.coordinator.data.get("mute"))
            if muted is not None:
                self._attr_is_volume_muted = muted

        if "sour" in self.coordinator.data:
            source = self.coordinator.data.get("sour")
            if source is not None:
                self._attr_source = self._get_source_translation_key(source)

        self.async_write_ha_state()

    async def async_turn_on(self) -> None:
        """Turn projector on."""
        if await self.coordinator.async_turn_on():
            self._attr_state = MediaPlayerState.ON
            self._attr_available = True
            self.async_write_ha_state()

    async def async_turn_off(self) -> None:
        """Turn projector off."""
        if await self.coordinator.async_turn_off():
            self._attr_state = MediaPlayerState.OFF
            self._attr_available = True
            self.async_write_ha_state()

    async def async_mute_volume(self, mute: bool) -> None:
        """Mute (true) or unmute (false) media player."""
        if mute:
            await self.coordinator.async_mute()
        else:
            await self.coordinator.async_unmute()

        muted = self._normalize_mute(self.coordinator.muted)
        if muted is not None:
            self._attr_is_volume_muted = muted
        self.async_write_ha_state()

    async def async_set_volume_level(self, volume: float) -> None:
        """Set volume level, range 0..1."""
        volume = int(volume * 20.0)
        if await self.coordinator.async_volume_level(volume):
            volume_level = self._normalize_volume_level(self.coordinator.volume)
            if volume_level is not None:
                self._attr_volume_level = volume_level
            self.async_write_ha_state()

    async def async_volume_up(self) -> None:
        """Increase volume."""
        if await self.coordinator.async_volume_up():
            volume_level = self._normalize_volume_level(self.coordinator.volume)
            if volume_level is not None:
                self._attr_volume_level = volume_level
            self._attr_is_volume_muted = False
            self.async_write_ha_state()

    async def async_volume_down(self) -> None:
        """Decrease volume."""
        if await self.coordinator.async_volume_down():
            volume_level = self._normalize_volume_level(self.coordinator.volume)
            if volume_level is not None:
                self._attr_volume_level = volume_level
            self._attr_is_volume_muted = False
            self.async_write_ha_state()

    async def async_select_source(self, source: str) -> None:
        """Set the input video source."""
        video_source = source.rstrip("1")
        if await self.coordinator.async_select_video_source(video_source):
            self._attr_source = source
            self.async_write_ha_state()
