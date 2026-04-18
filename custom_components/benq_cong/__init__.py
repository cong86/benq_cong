"""The BenQ Projector integration."""

from __future__ import annotations

import asyncio
import logging
import re
from datetime import timedelta
from time import monotonic
from typing import Any, Callable

import homeassistant.helpers.config_validation as cv
import serial
import voluptuous as vol
from benqprojector import BenQProjector, BenQProjectorSerial, BenQProjectorTelnet
from benqprojector.benqclasses import (
    BenQInvallidResponseError,
    BenQPromptTimeoutError,
    BenQResponseTimeoutError,
)
from homeassistant.config_entries import ConfigEntry
from homeassistant.const import (
    CONF_DEVICE_ID,
    CONF_HOST,
    CONF_PORT,
    CONF_TYPE,
    Platform,
)
from homeassistant.core import (
    CALLBACK_TYPE,
    HomeAssistant,
    ServiceCall,
    SupportsResponse,
    callback,
)
from homeassistant.exceptions import ConfigEntryNotReady
from homeassistant.helpers import entity_registry
from homeassistant.helpers.entity import DeviceInfo
from homeassistant.helpers.update_coordinator import DataUpdateCoordinator, UpdateFailed

from .const import (
    CONF_BAUD_RATE,
    CONF_DEFAULT_INTERVAL,
    CONF_INTERVAL,
    CONF_MODEL,
    CONF_SERIAL_PORT,
    CONF_TYPE_SERIAL,
    CONF_TYPE_TELNET,
    DOMAIN,
)
from .command_profile import OFFICIAL_UNSUPPORTED_COMMANDS, supports_command_by_profile

_LOGGER = logging.getLogger(__name__)
_BENQ_LOGGER = logging.getLogger("benqprojector.benqprojector")
_BENQ_CONNECTION_LOGGER = logging.getLogger("benqprojector.benqconnection")
_ASYNCIO_LOGGER = logging.getLogger("asyncio")

_BRIDGE_BLOCKED_COMMANDS = {"macaddr", "ltim2"}

# Official HT3550/W2700 class commands marked unsupported by RS232 table.
_W2700_UNSUPPORTED_COMMANDS = set(OFFICIAL_UNSUPPORTED_COMMANDS)


class _BridgeNoiseFilter(logging.Filter):
    """Filter noisy bridge warnings that are handled by retry/fallback logic."""

    SUPPRESSED = {
        "Timeout while waiting for response",
        "Failed to turn on projector, response: ?",
        "Failed to turn off projector, response: ?",
        "Unable to retrieve projector power state: Response timeout for command 'pow' and action '?'",
        "Failed to retrieve projector power state.",
    }

    def filter(self, record: logging.LogRecord) -> bool:
        message = record.getMessage()

        if re.match(r"^Problem communicating with\s+.+$", message):
            return False

        if message in self.SUPPRESSED:
            return False

        if re.match(
            r"^Projector model unexpectedly changed from Unknown to .+$", message
        ):
            return False

        match = re.match(r"^Command\s+([a-z0-9_]+)\s+not\s+supported$", message)
        if match and match.group(1).lower() in _W2700_UNSUPPORTED_COMMANDS:
            return False

        return True


class _AsyncioBridgeNoiseFilter(logging.Filter):
    """Suppress known benign asyncio socket noise from flaky TCP bridges."""

    SUPPRESSED = {
        "socket.send() raised exception.",
    }

    def filter(self, record: logging.LogRecord) -> bool:
        return record.getMessage() not in self.SUPPRESSED


class _BenQConnectionNoiseFilter(logging.Filter):
    """Suppress benign transport noise from flaky bridge framing."""

    SUPPRESSED = {
        "Incomplete read",
    }

    def filter(self, record: logging.LogRecord) -> bool:
        return record.getMessage() not in self.SUPPRESSED


def _install_bridge_noise_filter() -> None:
    """Install log filter once to suppress known false-positive bridge warnings."""
    if getattr(_BENQ_LOGGER, "_bridge_noise_filter_installed", False):
        return

    _BENQ_LOGGER.addFilter(_BridgeNoiseFilter())
    _BENQ_LOGGER._bridge_noise_filter_installed = True


def _install_asyncio_bridge_noise_filter() -> None:
    """Install asyncio noise filter once for known bridge-related socket errors."""
    if getattr(_ASYNCIO_LOGGER, "_bridge_noise_filter_installed", False):
        return

    _ASYNCIO_LOGGER.addFilter(_AsyncioBridgeNoiseFilter())
    _ASYNCIO_LOGGER._bridge_noise_filter_installed = True


def _install_benq_connection_noise_filter() -> None:
    """Install benqconnection noise filter once for known bridge transport quirks."""
    if getattr(_BENQ_CONNECTION_LOGGER, "_bridge_noise_filter_installed", False):
        return

    _BENQ_CONNECTION_LOGGER.addFilter(_BenQConnectionNoiseFilter())
    _BENQ_CONNECTION_LOGGER._bridge_noise_filter_installed = True


def _is_stray_query_echo(response: str, command) -> bool:
    """Return True when response looks like an echo for a different query command."""
    response_upper = str(response).strip().upper()
    match = re.match(r"^\*?([A-Z0-9]+)=\?#?$", response_upper)
    if not match:
        return False

    expected_command = getattr(command, "command", "")
    return bool(expected_command) and match.group(1).lower() != expected_command.lower()


def _clean_bridge_response(response: str) -> str:
    """Clean noisy bridge prefixes and keep the last valid BenQ frame."""
    cleaned = response.strip()
    cleaned = re.sub(
        r"(?:\+EVE)?Illegal\s+format\s*>*", "", cleaned, flags=re.IGNORECASE
    )
    cleaned = re.sub(r"^format\s*>+", "", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.lstrip(">")

    start = cleaned.rfind("*")
    end = cleaned.rfind("#")
    if start != -1 and end != -1 and end > start:
        cleaned = cleaned[start : end + 1]

    return cleaned or response


def _wrap_bridge_raw_command(command: str) -> str:
    """Wrap command in bridge-friendly framing with leading/trailing CRLF."""
    cmd = str(command or "").strip()
    if not cmd:
        return "\r\n"
    if not cmd.startswith("*"):
        cmd = f"*{cmd}"
    if not cmd.endswith("#"):
        cmd = f"{cmd}#"
    return f"\r\n{cmd}\r\n"


def _bridge_raw_command_variants(command: str) -> list[str]:
    """Return conservative framing variants with CRLF envelope preferred."""
    cmd = str(command or "").strip()
    if not cmd:
        return ["\r\n", "\n"]
    if not cmd.startswith("*"):
        cmd = f"*{cmd}"
    if not cmd.endswith("#"):
        cmd = f"{cmd}#"
    # DR164/162 has proven most stable with CRLF-wrapped frames.
    return [
        f"\r\n{cmd}\r\n",
        cmd,
    ]


def _parse_direct_query_value(command: str, payload: str) -> str | None:
    """Parse a command value from raw TCP query payload."""
    expected = str(command or "").strip().upper()
    for cmd_name, value in re.findall(r"\*?([A-Z0-9]+)=([^#\r\n]+)#", payload.upper()):
        if cmd_name == expected:
            return value.strip().lower()
    return None


def _extract_pow_state(response: str) -> str | None:
    """Extract pow state from bridge frames like '*POW=OFF#' or 'POW=OFF#'."""
    text = str(response).strip().upper()
    match = re.search(r"\*?POW=(ON|OFF)#?$", text)
    if not match:
        return None
    return match.group(1).lower()


def _extract_model_name(response: str) -> str | None:
    """Extract model name from bridge frame like '*MODELNAME=W2700i#'."""
    text = str(response).strip()
    match = re.search(r"\*?MODELNAME=([^#]+)#?$", text, re.IGNORECASE)
    if not match:
        return None
    return match.group(1).strip() or None


def _is_bridge_noise_response(response: str | None) -> bool:
    """Return True for known bridge garbage frames that should be ignored."""
    if response is None:
        return True

    text = str(response).strip()
    if not text:
        return True

    text_upper = text.upper()
    if text_upper in {"#", "*#", "+EVEILLEGAL", "ILLEGAL"}:
        return True

    if "ILLEGAL FORMAT" in text_upper:
        return True

    return False


def _patch_bridge_command_support(projector: BenQProjectorTelnet) -> None:
    """Disable optional commands that many bridges reject."""
    _orig_supports = projector.supports_command
    projector._bridge_runtime_unsupported_commands = set(_W2700_UNSUPPORTED_COMMANDS)

    def _supports_command(command):
        runtime_unsupported = getattr(projector, "_bridge_runtime_unsupported_commands", set())
        if isinstance(command, str) and command.lower() in _BRIDGE_BLOCKED_COMMANDS:
            return False
        if isinstance(command, str) and command.lower() in runtime_unsupported:
            return False
        if isinstance(command, str) and not supports_command_by_profile(command, "any"):
            return False
        return _orig_supports(command)

    projector.supports_command = _supports_command


def _sanitize_runtime_unsupported_commands(projector: BenQProjectorTelnet) -> None:
    """Keep runtime unsupported list aligned with official unsupported profile only."""
    runtime_unsupported = set(
        getattr(projector, "_bridge_runtime_unsupported_commands", set())
    )
    if not runtime_unsupported:
        return

    filtered = {
        command
        for command in runtime_unsupported
        if command in _W2700_UNSUPPORTED_COMMANDS
        or not supports_command_by_profile(command, "any")
    }

    if filtered != runtime_unsupported:
        projector._bridge_runtime_unsupported_commands = filtered
        _LOGGER.debug(
            "Sanitized runtime unsupported commands from %s to %s",
            sorted(runtime_unsupported),
            sorted(filtered),
        )


def _patch_bridge_response_parser(projector: BenQProjectorTelnet) -> None:
    """Patch parser to tolerate noisy bridge responses."""
    if getattr(projector, "_bridge_parser_patched", False):
        return

    _orig_parse = projector._parse_response

    def _parse_response(command, response, lowercase=True):
        command_name = getattr(command, "command", "")
        if isinstance(response, str):
            response = _clean_bridge_response(response)
            response_upper = str(response).upper()

            if _is_bridge_noise_response(response):
                return None

            # Some TCP-RS232 bridges drop the leading '*' from normal POW frames.
            if command_name == "pow":
                pow_state = _extract_pow_state(response)
                if pow_state is not None:
                    return pow_state if lowercase else pow_state.upper()

                # Bridge may interleave MODELNAME frame while we are querying power.
                if (model_name := _extract_model_name(response)) is not None:
                    projector.model = model_name
                    return None

            # Bridges may return unsupported item without spacing/canonical format.
            if "UNSUPPORTEDITEM" in response_upper or "UNSUPPORTED ITEM" in response_upper:
                if command_name:
                    command_key = command_name.lower()
                    if not supports_command_by_profile(command_key, "any"):
                        runtime_unsupported = getattr(
                            projector, "_bridge_runtime_unsupported_commands", set()
                        )
                        runtime_unsupported.add(command_key)
                        projector._bridge_runtime_unsupported_commands = runtime_unsupported
                _LOGGER.debug(
                    "Command reported unsupported by bridge, suppressing parser error for %s: %s",
                    command_name,
                    response,
                )
                if command_name == "modelname":
                    return projector.model or "Unknown"
                return None

        try:
            parsed = _orig_parse(command, response, lowercase)
            if parsed == "?" and getattr(command, "action", None) == "?":
                # Bridge sometimes answers '?' for unsupported/invalid query in current state.
                return None
            return parsed
        except BenQInvallidResponseError:
            # Bridge noise can intermittently inject unrelated query echoes.
            if _is_stray_query_echo(str(response), command):
                _LOGGER.debug(
                    "Ignoring stray query echo while parsing %s: %s",
                    getattr(command, "command", ""),
                    response,
                )
                return None

            response_upper = str(response).upper()
            if _is_bridge_noise_response(response):
                return None

            if getattr(command, "command", "") == "pow":
                if _extract_model_name(response) is not None:
                    return None

            if "FORMAT >" in response_upper or response_upper.startswith("NT="):
                _LOGGER.debug(
                    "Ignoring bridge format noise while parsing %s: %s",
                    getattr(command, "command", ""),
                    response,
                )
                return None

            if getattr(command, "command", "") == "modelname":
                if "*POW=" in response_upper or "UNSUPPORTEDITEM" in response_upper:
                    return projector.model or "Unknown"
            raise

    projector._parse_response = _parse_response
    projector._bridge_parser_patched = True


def _patch_bridge_echo_behavior(projector: BenQProjectorTelnet) -> None:
    """Disable strict command echo expectation for bridge connections."""
    if getattr(projector, "_bridge_echo_patched", False):
        return

    def _disable_echo_expectation() -> None:
        if hasattr(projector, "_expect_command_echo"):
            projector._expect_command_echo = False

        # Compatibility with older/newer implementations that keep parser state in a runner.
        runner = getattr(projector, "runner", None)
        if runner is not None:
            for attr in ("_expect_command_echo", "expect_command_echo"):
                if hasattr(runner, attr):
                    setattr(runner, attr, False)

    _disable_echo_expectation()

    if hasattr(projector, "_send_command"):
        _orig_send_command = projector._send_command
        projector._bridge_retry_in_progress = False
        projector._bridge_in_connect_attempt = getattr(
            projector, "_bridge_in_connect_attempt", False
        )

        async def _retry_after_reconnect(command, check_supported, lowercase_response):
            if getattr(projector, "_bridge_in_connect_attempt", False):
                return None

            if getattr(projector, "_bridge_retry_in_progress", False):
                return None

            projector._bridge_retry_in_progress = True
            try:
                await projector.disconnect()
                if await projector.connect():
                    _disable_echo_expectation()
                    return await _orig_send_command(
                        command, check_supported, lowercase_response
                    )
            finally:
                projector._bridge_retry_in_progress = False

            return None

        async def _send_command(command, check_supported=True, lowercase_response=True):
            _disable_echo_expectation()
            command_name = getattr(command, "command", None)
            action = getattr(command, "action", None)
            if command_name and not projector.supports_command(command_name):
                # Avoid upstream warning spam for unsupported optional commands.
                return None
            try:
                response = await _orig_send_command(
                    command, check_supported, lowercase_response
                )

                # Some bridges return cached non-power values while still in standby.
                # Only allow non-power responses to promote ON during POWERINGON.
                if (
                    command_name
                    and command_name != "pow"
                    and action == "?"
                    and response not in (None, "?")
                    and getattr(projector, "power_status", None)
                    == projector.POWERSTATUS_POWERINGON
                ):
                    projector.power_status = projector.POWERSTATUS_ON

                # Some bridges intermittently drop pow query responses.
                if (
                    command_name == "pow"
                    and action == "?"
                    and response in (None, "?")
                ):
                    raw_pow_response = await projector.send_raw_command("*pow=?#")
                    pow_state = _extract_pow_state(str(raw_pow_response))
                    if pow_state is not None:
                        projector.power_status = (
                            projector.POWERSTATUS_ON
                            if pow_state == "on"
                            else projector.POWERSTATUS_OFF
                        )
                        return pow_state

                    # Prefer conservative fallback to avoid false ON after shutdown.
                    if projector.power_status == projector.POWERSTATUS_OFF:
                        return "off"
                    if projector.power_status == projector.POWERSTATUS_POWERINGOFF:
                        return "off"

                    if getattr(projector, "_bridge_in_connect_attempt", False):
                        return response

                    _LOGGER.debug(
                        "Empty/unknown pow query response, reconnecting and retrying once"
                    )
                    retried = await _retry_after_reconnect(
                        command, check_supported, lowercase_response
                    )
                    if retried not in (None, "?"):
                        return retried

                    if projector.power_status == projector.POWERSTATUS_OFF:
                        return "off"
                    if projector.power_status == projector.POWERSTATUS_POWERINGOFF:
                        return "off"
                return response
            except BenQResponseTimeoutError:
                if getattr(projector, "_bridge_in_connect_attempt", False):
                    raise

                _LOGGER.debug(
                    "Response timeout for %s, reconnecting and retrying once",
                    getattr(command, "raw_command", command),
                )
                retried = await _retry_after_reconnect(
                    command, check_supported, lowercase_response
                )
                if retried is not None:
                    return retried
                raise
            except OSError as err:
                if getattr(projector, "_bridge_in_connect_attempt", False):
                    raise

                _LOGGER.debug(
                    "Socket send/read error for %s (%s), reconnecting and retrying once",
                    getattr(command, "raw_command", command),
                    err,
                )
                retried = await _retry_after_reconnect(
                    command, check_supported, lowercase_response
                )
                if retried is not None:
                    return retried
                raise

        projector._send_command = _send_command

    if hasattr(projector, "send_raw_command"):
        _orig_send_raw_command = projector.send_raw_command

        async def _send_raw_command(command: str):
            _disable_echo_expectation()
            last_response = None
            for framed in _bridge_raw_command_variants(command):
                try:
                    response = await _orig_send_raw_command(framed)
                    response_text = str(response or "").strip().upper()

                    # Keep trying other frame styles when bridge reports bad format/item.
                    if (
                        "ILLEGAL FORMAT" in response_text
                        or "UNSUPPORTED ITEM" in response_text
                        or "UNSUPPORTEDITEM" in response_text
                        or "BLOCK ITEM" in response_text
                    ):
                        last_response = response
                        continue

                    if response not in (None, "", "?"):
                        return response
                    last_response = response
                except Exception:  # noqa: BLE001
                    continue
            return last_response

        projector.send_raw_command = _send_raw_command

    projector._bridge_echo_patched = True


def _patch_bridge_prompt_fallback(projector: BenQProjectorTelnet) -> None:
    """Fallback to no-prompt mode when prompt waits are unstable on bridges."""
    if getattr(projector, "_bridge_prompt_fallback_patched", False):
        return

    if not hasattr(projector, "_wait_for_prompt"):
        return

    _orig_wait_for_prompt = projector._wait_for_prompt

    async def _wait_for_prompt():
        try:
            return await _orig_wait_for_prompt()
        except BenQPromptTimeoutError:
            _LOGGER.debug(
                "Prompt timeout on bridge %s, switching to no-prompt mode",
                projector.connection,
            )
            projector.has_prompt = False
            projector._has_to_wait_for_prompt = False
            return True

    projector._wait_for_prompt = _wait_for_prompt
    projector._bridge_prompt_fallback_patched = True


def _patch_bridge_raw_response_reader(projector: BenQProjectorTelnet) -> None:
    """Skip obviously unrelated bridge frames before parser validation."""
    if getattr(projector, "_bridge_raw_reader_patched", False):
        return

    if not hasattr(projector, "_read_raw_response"):
        return

    _orig_read_raw_response = projector._read_raw_response

    async def _read_raw_response(command):
        last_response = None
        for _ in range(6):
            response = await _orig_read_raw_response(command)
            if isinstance(response, str):
                response = _clean_bridge_response(response)
                response_upper = response.upper()

                if _is_bridge_noise_response(response):
                    _LOGGER.debug(
                        "Discarding bridge noise frame for %s: %s",
                        getattr(command, "command", ""),
                        response,
                    )
                    last_response = response
                    continue

                if getattr(command, "command", "") == "pow":
                    model_name = _extract_model_name(response)
                    if model_name is not None:
                        projector.model = model_name
                        _LOGGER.debug(
                            "Discarding interleaved MODELNAME frame during pow query: %s",
                            response,
                        )
                        last_response = response
                        continue

                if _is_stray_query_echo(response, command):
                    _LOGGER.debug(
                        "Discarding stray query echo for %s: %s",
                        getattr(command, "command", ""),
                        response,
                    )
                    last_response = response
                    continue

                if "FORMAT >" in response_upper or response_upper.startswith("NT="):
                    _LOGGER.debug(
                        "Discarding bridge format noise for %s: %s",
                        getattr(command, "command", ""),
                        response,
                    )
                    last_response = response
                    continue

            return response

        return last_response

    projector._read_raw_response = _read_raw_response
    projector._bridge_raw_reader_patched = True


def _patch_bridge_power_control(projector: BenQProjectorTelnet) -> None:
    """Use bridge-friendly power control to avoid false failures on '?' responses."""
    if getattr(projector, "_bridge_power_control_patched", False):
        return

    async def _confirm_power(target_on: bool) -> bool:
        for delay in (0.4, 0.8, 1.2):
            await asyncio.sleep(delay)
            if not await projector.update_power():
                continue

            if target_on and projector.power_status in [
                projector.POWERSTATUS_POWERINGON,
                projector.POWERSTATUS_ON,
            ]:
                return True

            if not target_on and projector.power_status in [
                projector.POWERSTATUS_POWERINGOFF,
                projector.POWERSTATUS_OFF,
            ]:
                return True

        return False

    async def _bridge_turn_on() -> bool:
        state = await projector.send_command("pow")
        if state == "on":
            projector.power_status = projector.POWERSTATUS_ON
            return True

        response = await projector.send_command("pow", "on")
        if response == "on":
            projector.power_status = projector.POWERSTATUS_POWERINGON
            return True

        # Many bridges answer '?' for accepted state-changing commands.
        if response in ("?", None, False):
            await projector.send_raw_command("*pow=on#")
            projector.power_status = projector.POWERSTATUS_POWERINGON
            return await _confirm_power(target_on=True)

        return await _confirm_power(target_on=True)

    async def _bridge_turn_off() -> bool:
        state = await projector.send_command("pow")
        if state == "off":
            projector.power_status = projector.POWERSTATUS_OFF
            return True

        response = await projector.send_command("pow", "off")
        if response == "off":
            projector.power_status = projector.POWERSTATUS_POWERINGOFF
            return True

        if response in ("?", None, False):
            await projector.send_raw_command("*pow=off#")
            projector.power_status = projector.POWERSTATUS_POWERINGOFF
            return await _confirm_power(target_on=False)

        return await _confirm_power(target_on=False)

    projector.turn_on = _bridge_turn_on
    projector.turn_off = _bridge_turn_off
    projector._bridge_power_control_patched = True


async def _is_tcp_port_reachable(host: str, port: int, timeout: float = 3.0) -> bool:
    """Return true when host:port is reachable at TCP level."""
    try:
        _, writer = await asyncio.wait_for(
            asyncio.open_connection(host, port), timeout=timeout
        )
        writer.close()
        await writer.wait_closed()
        return True
    except Exception:
        return False

PLATFORMS: list[Platform] = [
    Platform.MEDIA_PLAYER,
    Platform.SENSOR,
    Platform.SWITCH,
    Platform.SELECT,
    Platform.NUMBER,
    Platform.BUTTON,
]

CONF_SERVICE_COMMAND = "command"
CONF_SERVICE_ACTION = "action"

SERVICE_SEND_SCHEMA = vol.Schema(
    {
        vol.Required(CONF_DEVICE_ID): cv.string,
        vol.Required(CONF_SERVICE_COMMAND): cv.string,
        vol.Optional(CONF_SERVICE_ACTION): cv.string,
    }
)
SERVICE_SEND_RAW_SCHEMA = vol.Schema(
    {
        vol.Required(CONF_DEVICE_ID): cv.string,
        vol.Required(CONF_SERVICE_COMMAND): cv.string,
    }
)


class BenQProjectorCoordinator(DataUpdateCoordinator):
    """BenQ Projector Data Update Coordinator."""

    unique_id = None
    model = None
    device_info: DeviceInfo = None
    _OFF_SAFE_READ_COMMANDS = {
        "pow",
    }

    def __init__(self, hass, projector: BenQProjector, poll_interval: int = 5):
        """Initialize BenQ Projector Data Update Coordinator."""
        super().__init__(
            hass,
            _LOGGER,
            # Name of the data. For logging purposes.
            name=__name__,
        )

        self.projector = projector
        self.projector.add_listener(self._listener)
        self._last_non_power_success = 0.0
        self._command_lock = asyncio.Lock()
        self._post_power_on_task: asyncio.Task | None = None
        self._state_poll_task: asyncio.Task | None = None
        self._state_poll_index = 0
        self._poll_interval = max(3, int(poll_interval or 5))
        self._power_on_grace_until = 0.0
        self._query_cache: dict[str, tuple[float, Any]] = {}
        self._query_min_interval: dict[str, float] = {
            "pow": 1.5,
            "sour": 2.0,
            "vol": 2.0,
            "mute": 2.0,
            "modelname": 60.0,
            "ltim": 30.0,
            "appmod": 10.0,
            "asp": 10.0,
            "ct": 10.0,
            "lampm": 10.0,
            "3d": 10.0,
            "bc": 10.0,
            "pp": 10.0,
            "blank": 4.0,
            "freeze": 4.0,
            "highaltitude": 30.0,
            "directpower": 30.0,
        }

        self.unique_id = self.projector.unique_id
        model = self.projector.model
        if model is not None:
            model = model

        self.device_info = DeviceInfo(
            identifiers={(DOMAIN, self.unique_id)},
            name=f"BenQ {model}",
            model=model,
            manufacturer="BenQ",
        )

    @property
    def power_status(self):
        status = self.projector.power_status
        if (
            status
            in [
                self.projector.POWERSTATUS_UNKNOWN,
                self.projector.POWERSTATUS_OFF,
                self.projector.POWERSTATUS_POWERINGOFF,
            ]
            and self._is_in_power_on_grace()
        ):
            return self.projector.POWERSTATUS_POWERINGON

        return status

    def _is_in_power_on_grace(self) -> bool:
        return monotonic() < self._power_on_grace_until

    def _start_power_on_grace(self, seconds: float = 45.0) -> None:
        self._power_on_grace_until = monotonic() + seconds

    def _clear_power_on_grace(self) -> None:
        self._power_on_grace_until = 0.0

    def _is_off_safe_read_command(self, command: str) -> bool:
        return str(command or "").lower() in self._OFF_SAFE_READ_COMMANDS

    @property
    def volume(self):
        return self.projector.volume

    @property
    def muted(self):
        return self.projector.muted

    @property
    def video_source(self):
        return self.projector.video_source

    @property
    def video_sources(self):
        return self.projector.video_sources

    @callback
    def _listener(self, command: str, data):
        clear_non_power_state = False

        if command == "pow":
            data_text = str(data).strip().lower() if data is not None else ""

            if data_text == "on":
                self.projector.power_status = self.projector.POWERSTATUS_ON
                self._clear_power_on_grace()
            elif data_text == "off" and self._is_in_power_on_grace():
                # Bridge can report stale OFF right after accepted power-on command.
                self.projector.power_status = self.projector.POWERSTATUS_POWERINGON
                data = "on"
            elif data_text == "off":
                self.projector.power_status = self.projector.POWERSTATUS_OFF
                # Stop ON inference window immediately after explicit power-off.
                self._last_non_power_success = 0.0
                self._clear_power_on_grace()
                clear_non_power_state = True

        if command != "pow" and data not in (None, "", "?", "unknown", "unavailable"):
            self._last_non_power_success = monotonic()

        if data not in (None, "", "?", "unknown", "unavailable"):
            self._query_cache[command] = (monotonic(), data)

        merged = dict(self.data or {})
        if clear_non_power_state:
            merged = {
                key: value
                for key, value in merged.items()
                if self._is_off_safe_read_command(key)
            }
            self._query_cache = {
                key: value
                for key, value in self._query_cache.items()
                if self._is_off_safe_read_command(key)
            }
        merged[command] = data
        self.async_set_updated_data(merged)

    # async def async_connect(self):
    #     try:
    #         if not await self.projector.connect():
    #             raise ConfigEntryNotReady(
    #                 f"Unable to connect to BenQ projector on {self.projector.connection}"
    #             )
    #     except TimeoutError as ex:
    #         raise ConfigEntryNotReady(
    #             f"Unable to connect to BenQ projector on {self.projector.connection}", ex
    #         )
    #
    #     _LOGGER.debug("Connected to BenQ projector on %s", self.projector.connection)
    #
    #     self.unique_id = self.projector.unique_id
    #     model = self.projector.model
    #
    #     self.device_info = DeviceInfo(
    #         identifiers={(DOMAIN, self.unique_id)},
    #         name=f"BenQ {model}",
    #         model=model,
    #         manufacturer="BenQ",
    #     )

    async def async_disconnect(self):
        if self._state_poll_task and not self._state_poll_task.done():
            self._state_poll_task.cancel()
            self._state_poll_task = None

        if self._post_power_on_task and not self._post_power_on_task.done():
            self._post_power_on_task.cancel()
            self._post_power_on_task = None

        await self.projector.disconnect()
        _LOGGER.debug(
            "Disconnected from BenQ projector on %s", self.projector.connection
        )

    @callback
    def async_add_listener(
        self, update_callback: CALLBACK_TYPE, context: Any = None
    ) -> Callable[[], None]:
        self.projector.add_listener(command=context)

        return super().async_add_listener(update_callback, context)

    def supports_command(self, command: str, operation: str = "any"):
        if not supports_command_by_profile(command, operation):
            return False
        return self.projector.supports_command(command)

    async def _async_retry_operation(
        self, operation_name: str, operation, retry_on_failure: bool = True
    ):
        """Retry once after reconnect when a command/control operation fails."""
        result = await operation()
        if result not in (None, False):
            return result

        if not retry_on_failure:
            return result

        retry_message = "%s failed on %s, reconnecting and retrying once"
        if operation_name in {"send_raw_command"}:
            _LOGGER.debug(retry_message, operation_name, self.projector.connection)
        else:
            _LOGGER.warning(retry_message, operation_name, self.projector.connection)

        await self.projector.disconnect()
        if not await self.projector.connect():
            return result

        return await operation()

    async def _async_confirm_power_state(self, target_on: bool) -> bool:
        """Confirm power transition to tolerate noisy bridge acknowledgements."""
        for delay in (0.4, 0.8, 1.2, 2.0, 3.0, 4.0):
            await asyncio.sleep(delay)
            try:
                if not await self.projector.update_power():
                    continue
            except Exception:  # noqa: BLE001
                continue

            if target_on and self.projector.power_status in [
                self.projector.POWERSTATUS_POWERINGON,
                self.projector.POWERSTATUS_ON,
            ]:
                return True

            if not target_on and self.projector.power_status in [
                self.projector.POWERSTATUS_POWERINGOFF,
                self.projector.POWERSTATUS_OFF,
            ]:
                return True

        return False

    async def _async_get_power_status_fresh(self) -> int:
        """Query current power status with cache bypass and synchronize state."""
        self._query_cache.pop("pow", None)

        response = None
        try:
            response = await self.async_send_command("pow")
        except Exception:  # noqa: BLE001
            response = None

        response_text = str(response).strip().lower() if response is not None else ""
        if response_text == "on":
            self._listener("pow", "on")
            return self.projector.POWERSTATUS_ON

        if response_text == "off":
            self._listener("pow", "off")
            return self.projector.POWERSTATUS_OFF

        return self.projector.power_status

    async def _async_send_power_off_double_sequence(self) -> None:
        """Send OFF command twice to match projector confirm-off behavior."""
        for _ in range(2):
            try:
                await self.async_send_command("pow", "off")
            except Exception:  # noqa: BLE001
                pass

            try:
                await self.async_send_raw_command("*pow=off#")
            except Exception:  # noqa: BLE001
                pass

            await asyncio.sleep(0.15)

    async def _async_post_power_on_refresh(self) -> None:
        """Warm up key states after power-on when bridge polling is delayed/noisy."""
        command_groups = [
            ("pow",),
            ("pow", "sour", "vol", "mute", "con", "bri"),
            (
                "pow",
                "sour",
                "vol",
                "mute",
                "con",
                "bri",
                "color",
                "sharp",
                "appmod",
                "asp",
                "lampm",
                "ct",
                "3d",
                "pp",
                "bc",
                "blank",
                "freeze",
            ),
        ]

        for index, commands in enumerate(command_groups):
            if index == 0:
                await asyncio.sleep(2.0)
            elif index == 1:
                await asyncio.sleep(4.0)
            else:
                await asyncio.sleep(6.0)

            for command in commands:
                if not supports_command_by_profile(command, "read"):
                    continue

                # When projector is already OFF, avoid commands that are known to return
                # Illegal format in standby and only keep safe status probes.
                if self.projector.power_status in [
                    self.projector.POWERSTATUS_OFF,
                    self.projector.POWERSTATUS_POWERINGOFF,
                ] and not self._is_off_safe_read_command(command):
                    continue

                try:
                    response = await self.async_send_command(command)
                except Exception:  # noqa: BLE001
                    continue

                if response in (None, False, "", "?", "unknown", "unavailable"):
                    continue

                self._listener(command, response)

            # Only stop early after the full warmup round that includes select-related commands.
            if index >= 2 and self.power_status == self.projector.POWERSTATUS_ON:
                return

    def _get_state_poll_commands(self) -> tuple[str, ...]:
        """Return a conservative read command set for periodic state maintenance."""
        if self.power_status in [
            self.projector.POWERSTATUS_OFF,
            self.projector.POWERSTATUS_POWERINGOFF,
            self.projector.POWERSTATUS_UNKNOWN,
        ]:
            # Power-gated polling: only poll power while OFF/UNKNOWN.
            return ("pow",)

        if self.power_status == self.projector.POWERSTATUS_POWERINGON:
            return ("pow", "sour")

        groups: tuple[tuple[str, ...], ...] = (
            ("pow", "sour", "blank"),
            ("pow", "modelname", "ltim"),
            ("pow", "vol", "mute"),
            ("pow", "con", "bri"),
            ("pow", "color", "sharp"),
            ("pow", "appmod", "asp"),
            ("pow", "lampm", "ct"),
            ("pow", "3d", "pp", "bc", "freeze"),
        )
        commands = groups[self._state_poll_index % len(groups)]
        self._state_poll_index += 1
        return commands

    async def _async_poll_state_once(self) -> None:
        """Poll one low-frequency state batch to keep HA entities in sync."""
        invalid_values = (None, False, "", "?", "unknown", "unavailable")

        if hasattr(self.projector, "_bridge_runtime_unsupported_commands"):
            _sanitize_runtime_unsupported_commands(self.projector)

        for command in self._get_state_poll_commands():
            if not supports_command_by_profile(command, "read"):
                continue

            try:
                result = await self.async_send_command(command)
            except Exception:  # noqa: BLE001
                continue

            if result in invalid_values:
                continue

            self._listener(command, result)
            await asyncio.sleep(0.20)

    async def _async_state_poll_loop(self) -> None:
        """Background polling loop for bridges with lossy/partial unsolicited updates."""
        try:
            while True:
                try:
                    await self._async_poll_state_once()
                except Exception:  # noqa: BLE001
                    pass
                if self.power_status in [
                    self.projector.POWERSTATUS_OFF,
                    self.projector.POWERSTATUS_POWERINGOFF,
                    self.projector.POWERSTATUS_UNKNOWN,
                ]:
                    sleep_for = max(self._poll_interval, 8)
                elif self.power_status == self.projector.POWERSTATUS_POWERINGON:
                    sleep_for = max(self._poll_interval, 5)
                else:
                    sleep_for = max(self._poll_interval, 6)

                await asyncio.sleep(sleep_for)
        except asyncio.CancelledError:
            return

    def async_start_state_polling(self) -> None:
        """Start periodic state polling task if not already running."""
        if self._state_poll_task and not self._state_poll_task.done():
            return

        if hasattr(self.hass, "async_create_background_task"):
            try:
                self._state_poll_task = self.hass.async_create_background_task(
                    self._async_state_poll_loop(),
                    "benqprojector_state_poll_loop",
                )
                return
            except TypeError:
                # Compatibility with HA versions that require eager_start argument.
                self._state_poll_task = self.hass.async_create_background_task(
                    self._async_state_poll_loop(),
                    "benqprojector_state_poll_loop",
                    eager_start=True,
                )
                return

        # Fallback for older HA versions.
        self._state_poll_task = asyncio.create_task(
            self._async_state_poll_loop(), name="benqprojector_state_poll_loop"
        )

    async def _async_reconnect_once(self) -> bool:
        await self.projector.disconnect()
        return await self.projector.connect()

    @callback
    def _publish_state_snapshot(self) -> None:
        """Notify entities to recompute availability from current power_status."""
        self.async_set_updated_data(dict(self.data or {}))

    @callback
    def _set_power_status(self, status: int, *, reset_non_power: bool = False) -> None:
        """Update cached power status and broadcast coordinator update."""
        self.projector.power_status = status

        if status in [
            self.projector.POWERSTATUS_POWERINGON,
            self.projector.POWERSTATUS_ON,
        ] and hasattr(self.projector, "_bridge_runtime_unsupported_commands"):
            _sanitize_runtime_unsupported_commands(self.projector)

        if status in [
            self.projector.POWERSTATUS_OFF,
            self.projector.POWERSTATUS_POWERINGOFF,
        ]:
            self._clear_power_on_grace()

        if reset_non_power:
            self._last_non_power_success = 0.0
        self._publish_state_snapshot()

    def _get_connection_target(self) -> tuple[str, int] | tuple[None, None]:
        """Resolve host/port for direct TCP fallback sender."""
        host = getattr(self.projector, "host", None)
        port = getattr(self.projector, "port", None)
        if host and port:
            return str(host), int(port)

        connection = str(getattr(self.projector, "connection", ""))
        if ":" in connection:
            host_part, port_part = connection.rsplit(":", 1)
            try:
                return host_part, int(port_part)
            except ValueError:
                return None, None

        return None, None

    async def _async_send_power_tcp_direct(self, target_on: bool) -> bool:
        """Send power frame directly over TCP, bypassing benqprojector stack."""
        host, port = self._get_connection_target()
        if not host or not port:
            return False

        base_with_star = "*pow=on#" if target_on else "*pow=off#"
        frames = _bridge_raw_command_variants(base_with_star)

        try:
            reader, writer = await asyncio.wait_for(
                asyncio.open_connection(host, port), timeout=1.5
            )

            for _round in range(2):
                for frame in frames:
                    writer.write(frame.encode("ascii", errors="ignore"))
                    await writer.drain()
                    await asyncio.sleep(0.12)
                await asyncio.sleep(0.15)

            await asyncio.sleep(0.20)

            writer.close()
            await writer.wait_closed()
            _LOGGER.debug(
                "Emergency direct TCP power command sent: %s on %s:%s",
                base_with_star,
                host,
                port,
            )
            return True
        except Exception as err:  # noqa: BLE001
            _LOGGER.warning(
                "Emergency direct TCP power command failed (%s) on %s:%s: %s",
                base_with_star,
                host,
                port,
                err,
            )
            return False

    async def _async_query_tcp_direct(self, command: str) -> str | None:
        """Query command value via direct TCP channel, bypassing library parser stack."""
        host, port = self._get_connection_target()
        if not host or not port:
            return None

        query_with_star = f"*{command}=?#"
        query_no_star = f"{command}=?#"
        query_frames = [
            _wrap_bridge_raw_command(query_with_star),
            _wrap_bridge_raw_command(query_no_star),
        ]
        payload = ""

        try:
            reader, writer = await asyncio.wait_for(
                asyncio.open_connection(host, port), timeout=1.5
            )

            for frame_index, frame in enumerate(query_frames):
                writer.write(frame.encode("ascii", errors="ignore"))
                await writer.drain()
                await asyncio.sleep(0.1)

                # Give the preferred CRLF frame more time before trying fallback.
                read_attempts = 6 if frame_index == 0 else 3
                for _ in range(read_attempts):
                    try:
                        chunk = await asyncio.wait_for(reader.read(256), timeout=0.25)
                    except asyncio.TimeoutError:
                        break
                    if not chunk:
                        break
                    payload += chunk.decode("ascii", errors="ignore")

                    if _parse_direct_query_value(command, payload) is not None:
                        break

                if _parse_direct_query_value(command, payload) is not None:
                    break

            writer.close()
            await writer.wait_closed()
        except Exception as err:  # noqa: BLE001
            _LOGGER.debug(
                "Direct TCP query failed for %s on %s:%s: %s",
                command,
                host,
                port,
                err,
            )
            return None

        parsed = _parse_direct_query_value(command, payload)
        if parsed is not None:
            return parsed

        cleaned = _clean_bridge_response(payload)
        return _parse_direct_query_value(command, cleaned)

    async def _async_send_power_raw_emergency(self, target_on: bool) -> bool:
        """Send power command via direct raw channel, bypassing coordinator lock."""
        raw_command = "*pow=on#" if target_on else "*pow=off#"
        _LOGGER.info(
            "Power button pressed: sending emergency raw command %s on %s",
            raw_command,
            self.projector.connection,
        )

        tcp_direct_sent = await self._async_send_power_tcp_direct(target_on=target_on)
        if tcp_direct_sent:
            _LOGGER.debug(
                "Direct TCP power frame sent for %s, continuing raw fallback for reliability",
                raw_command,
            )

        async def _send_once() -> bool:
            try:
                await asyncio.wait_for(
                    self.projector.send_raw_command(raw_command), timeout=0.8
                )
                return True
            except asyncio.TimeoutError:
                # Bridges often execute the command but never deliver a parseable response.
                _LOGGER.debug(
                    "Emergency raw power command timed out (treated as sent): %s on %s",
                    raw_command,
                    self.projector.connection,
                )
                return True

        try:
            await _send_once()
            # Some bridges only accept the command intermittently; send once more.
            await asyncio.sleep(0.15)
            await _send_once()
            _LOGGER.debug(
                "Emergency raw power command sent: %s on %s",
                raw_command,
                self.projector.connection,
            )
            return True
        except Exception as err:  # noqa: BLE001
            _LOGGER.warning(
                "Emergency raw power command failed (%s) on %s: %s. Reconnecting and retrying once",
                raw_command,
                self.projector.connection,
                err,
            )
            try:
                await self.projector.disconnect()
                if await self.projector.connect():
                    await _send_once()
                    _LOGGER.debug(
                        "Emergency raw power command sent after reconnect: %s on %s",
                        raw_command,
                        self.projector.connection,
                    )
                    return True
            except Exception as retry_err:  # noqa: BLE001
                _LOGGER.warning(
                    "Emergency raw power command retry failed (%s) on %s: %s",
                    raw_command,
                    self.projector.connection,
                    retry_err,
                )

            return False

    async def _async_reinforce_power_command(self, target_on: bool) -> None:
        """Reinforce power command when first send was accepted but not executed."""
        action = "on" if target_on else "off"
        raw_command = "*pow=on#" if target_on else "*pow=off#"

        for _ in range(2):
            try:
                await self.async_send_command("pow", action)
            except Exception:  # noqa: BLE001
                pass

            try:
                await self.async_send_raw_command(raw_command)
            except Exception:  # noqa: BLE001
                pass

            await asyncio.sleep(0.15)

    async def async_send_command(self, command: str, action: str | None = None):
        invalid_values = (None, False, "", "?", "unknown", "unavailable")

        if action is None:
            now = monotonic()
            cache = self._query_cache.get(command)
            min_interval = self._query_min_interval.get(command, 5.0)
            if cache and (now - cache[0]) < min_interval:
                return cache[1]

        async with self._command_lock:
            if action is None:
                now = monotonic()
                cache = self._query_cache.get(command)
                min_interval = self._query_min_interval.get(command, 5.0)
                if cache and (now - cache[0]) < min_interval:
                    return cache[1]

            if action:
                result = await self._async_retry_operation(
                    f"send_command({command}={action})",
                    lambda: self.projector.send_command(command, action),
                )

                # Bridge can execute write commands but return empty/unknown parsed response.
                # Retry once using raw command frame for better interoperability.
                if result in (None, False, "?"):
                    raw_command = f"*{command}={action}#"
                    raw_result = await self._async_retry_operation(
                        f"send_raw_fallback({command}={action})",
                        lambda: self.projector.send_raw_command(raw_command),
                    )
                    if raw_result is not None:
                        result = action

                # Invalidate query cache after write attempts so next read gets fresh state.
                self._query_cache.pop(command, None)
                return result

            result = await self._async_retry_operation(
                f"send_command({command})",
                lambda: self.projector.send_command(command),
                retry_on_failure=False,
            )

            if result in invalid_values:
                if self.projector.power_status in [
                    self.projector.POWERSTATUS_OFF,
                    self.projector.POWERSTATUS_POWERINGOFF,
                ] and not self._is_off_safe_read_command(command):
                    # Skip duplicate direct query for OFF-unsafe commands.
                    pass
                else:
                    direct_result = await self._async_query_tcp_direct(command)
                    if direct_result not in invalid_values:
                        result = direct_result

            if result not in invalid_values:
                self._query_cache[command] = (monotonic(), result)

            return result

    async def async_send_raw_command(self, command: str):
        async with self._command_lock:
            return await self._async_retry_operation(
                "send_raw_command", lambda: self.projector.send_raw_command(command)
            )

    async def async_turn_on(self) -> bool:
        _LOGGER.info("Processing turn_on request for %s", self.projector.connection)

        current_power_status = await self._async_get_power_status_fresh()
        if current_power_status in [
            self.projector.POWERSTATUS_POWERINGON,
            self.projector.POWERSTATUS_ON,
        ]:
            self._set_power_status(self.projector.POWERSTATUS_ON)
            _LOGGER.info(
                "turn_on skipped on %s because projector is already on",
                self.projector.connection,
            )
            return True

        self._start_power_on_grace()
        if await self._async_send_power_raw_emergency(target_on=True):
            self._set_power_status(self.projector.POWERSTATUS_POWERINGON)
            if await self._async_confirm_power_state(target_on=True):
                _LOGGER.info("turn_on confirmed via power state on %s", self.projector.connection)
                if self._post_power_on_task and not self._post_power_on_task.done():
                    self._post_power_on_task.cancel()
                self._post_power_on_task = self.hass.async_create_task(
                    self._async_post_power_on_refresh()
                )
                return True
            _LOGGER.warning(
                "turn_on not confirmed after emergency send on %s, reinforcing power command",
                self.projector.connection,
            )
            await self._async_reinforce_power_command(target_on=True)
            await self._async_confirm_power_state(target_on=True)
            if self._post_power_on_task and not self._post_power_on_task.done():
                self._post_power_on_task.cancel()
            self._post_power_on_task = self.hass.async_create_task(
                self._async_post_power_on_refresh()
            )
            return True

        result = await self.async_send_command("pow", "on")
        if result in ("on", "?", "+"):
            self._set_power_status(self.projector.POWERSTATUS_POWERINGON)
            confirmed = await self._async_confirm_power_state(target_on=True)
            if self._post_power_on_task and not self._post_power_on_task.done():
                self._post_power_on_task.cancel()
            self._post_power_on_task = self.hass.async_create_task(
                self._async_post_power_on_refresh()
            )
            return confirmed or True

        try:
            await self.async_send_raw_command("*pow=on#")
        except Exception:  # noqa: BLE001
            pass
        self._set_power_status(self.projector.POWERSTATUS_POWERINGON)
        if await self._async_confirm_power_state(target_on=True):
            return True

        _LOGGER.warning(
            "turn_on failed on %s, reconnecting and retrying once",
            self.projector.connection,
        )
        if not await self._async_reconnect_once():
            return False

        result = await self.async_send_command("pow", "on")
        if result in ("on", "?", "+"):
            self._set_power_status(self.projector.POWERSTATUS_POWERINGON)

        if await self._async_confirm_power_state(target_on=True):
            return True

        try:
            await self.async_send_raw_command("*pow=on#")
        except Exception:  # noqa: BLE001
            pass
        self._set_power_status(self.projector.POWERSTATUS_POWERINGON)
        confirmed = await self._async_confirm_power_state(target_on=True)
        if self._post_power_on_task and not self._post_power_on_task.done():
            self._post_power_on_task.cancel()
        self._post_power_on_task = self.hass.async_create_task(
            self._async_post_power_on_refresh()
        )
        return confirmed or True

        return False

    async def async_turn_off(self) -> bool:
        _LOGGER.info("Processing turn_off request for %s", self.projector.connection)

        current_power_status = await self._async_get_power_status_fresh()
        if current_power_status in [
            self.projector.POWERSTATUS_OFF,
            self.projector.POWERSTATUS_POWERINGOFF,
        ]:
            self._set_power_status(self.projector.POWERSTATUS_OFF, reset_non_power=True)
            _LOGGER.info(
                "turn_off skipped on %s because projector is already off",
                self.projector.connection,
            )
            return True

        self._clear_power_on_grace()
        if self._post_power_on_task and not self._post_power_on_task.done():
            self._post_power_on_task.cancel()

        await self._async_send_power_off_double_sequence()
        if await self._async_confirm_power_state(target_on=False):
            self._set_power_status(self.projector.POWERSTATUS_OFF, reset_non_power=True)
            return True

        if await self._async_send_power_raw_emergency(target_on=False):
            self._set_power_status(self.projector.POWERSTATUS_POWERINGOFF, reset_non_power=True)
            if await self._async_confirm_power_state(target_on=False):
                self._set_power_status(self.projector.POWERSTATUS_OFF, reset_non_power=True)
                _LOGGER.info("turn_off confirmed via power state on %s", self.projector.connection)
                return True
            _LOGGER.warning(
                "turn_off not confirmed after emergency send on %s, reinforcing power command",
                self.projector.connection,
            )
            await self._async_reinforce_power_command(target_on=False)
            if await self._async_confirm_power_state(target_on=False):
                self._set_power_status(self.projector.POWERSTATUS_OFF, reset_non_power=True)
            return True

        result = await self.async_send_command("pow", "off")
        if result in ("off", "?", "+"):
            self._set_power_status(self.projector.POWERSTATUS_POWERINGOFF, reset_non_power=True)
            confirmed = await self._async_confirm_power_state(target_on=False)
            if confirmed:
                self._set_power_status(self.projector.POWERSTATUS_OFF, reset_non_power=True)
            return confirmed

        try:
            await self.async_send_raw_command("*pow=off#")
        except Exception:  # noqa: BLE001
            pass
        self._set_power_status(self.projector.POWERSTATUS_POWERINGOFF, reset_non_power=True)
        if await self._async_confirm_power_state(target_on=False):
            self._set_power_status(self.projector.POWERSTATUS_OFF, reset_non_power=True)
            return True

        _LOGGER.warning(
            "turn_off failed on %s, reconnecting and retrying once",
            self.projector.connection,
        )
        if not await self._async_reconnect_once():
            return False

        result = await self.async_send_command("pow", "off")
        if result in ("off", "?", "+"):
            self._set_power_status(self.projector.POWERSTATUS_POWERINGOFF, reset_non_power=True)

        if await self._async_confirm_power_state(target_on=False):
            self._set_power_status(self.projector.POWERSTATUS_OFF, reset_non_power=True)
            return True

        try:
            await self.async_send_raw_command("*pow=off#")
        except Exception:  # noqa: BLE001
            pass
        self._set_power_status(self.projector.POWERSTATUS_POWERINGOFF, reset_non_power=True)
        confirmed = await self._async_confirm_power_state(target_on=False)
        if confirmed:
            self._set_power_status(self.projector.POWERSTATUS_OFF, reset_non_power=True)
        return confirmed

        return False

    async def async_mute(self) -> bool:
        async with self._command_lock:
            return await self.projector.mute()

    async def async_unmute(self) -> bool:
        async with self._command_lock:
            return await self.projector.unmute()

    async def async_volume_level(self, volume: int):
        async with self._command_lock:
            return await self.projector.volume_level(volume)

    async def async_volume_up(self):
        async with self._command_lock:
            return await self.projector.volume_up()

    async def async_volume_down(self):
        async with self._command_lock:
            return await self.projector.volume_down()

    async def async_select_video_source(self, source: str):
        async with self._command_lock:
            return await self.projector.select_video_source(source)


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Set up BenQ Projector from a config entry."""
    projector = None

    model = entry.data.get(CONF_MODEL)
    if isinstance(model, str) and model.lower() == "unknown":
        model = None
    conf_type = entry.data.get(CONF_TYPE, CONF_TYPE_SERIAL)
    interval = entry.options.get(CONF_INTERVAL, CONF_DEFAULT_INTERVAL)

    if conf_type == CONF_TYPE_TELNET:
        host = entry.data[CONF_HOST]
        port = entry.data[CONF_PORT]
        telnet_model = model or "Unknown"
        _install_bridge_noise_filter()
        _install_asyncio_bridge_noise_filter()
        _install_benq_connection_noise_filter()

        # Try both prompt modes; TCP-to-RS232 bridges differ by model/firmware.
        for has_prompt in (False, True):
            candidate = BenQProjectorTelnet(
                host, port, telnet_model, has_prompt=has_prompt
            )
            _patch_bridge_command_support(candidate)
            _patch_bridge_response_parser(candidate)
            _patch_bridge_echo_behavior(candidate)
            _patch_bridge_prompt_fallback(candidate)
            _patch_bridge_raw_response_reader(candidate)
            _patch_bridge_power_control(candidate)

            candidate._bridge_in_connect_attempt = True
            try:
                if await candidate.connect(interval=interval):
                    projector = candidate
                    _LOGGER.info(
                        "Connected to %s:%s using has_prompt=%s",
                        host,
                        port,
                        has_prompt,
                    )
                    break
            finally:
                candidate._bridge_in_connect_attempt = False

            await candidate.disconnect()

        if projector is None:
            # Bridge startup can fail handshake intermittently while TCP is reachable.
            # Keep integration loaded in degraded mode and let command paths reconnect.
            if await _is_tcp_port_reachable(host, port):
                projector = BenQProjectorTelnet(
                    host, port, telnet_model, has_prompt=False
                )
                _patch_bridge_command_support(projector)
                _patch_bridge_response_parser(projector)
                _patch_bridge_echo_behavior(projector)
                _patch_bridge_prompt_fallback(projector)
                _patch_bridge_raw_response_reader(projector)
                _patch_bridge_power_control(projector)
                _LOGGER.warning(
                    "Proceeding with degraded startup for %s:%s; initial handshake failed",
                    host,
                    port,
                )
            else:
                raise ConfigEntryNotReady(f"Unable to connect to device {host}:{port}")
    else:
        serial_port = entry.data[CONF_SERIAL_PORT]
        baud_rate = entry.data[CONF_BAUD_RATE]

        projector = BenQProjectorSerial(serial_port, baud_rate, model)

    @callback
    def _async_migrate_entity_entry(
        registry_entry: entity_registry.RegistryEntry,
    ) -> dict[str, Any] | None:
        """
        Migrates old unique ID to the new unique ID.
        """
        if registry_entry.entity_id.startswith(
            "media_player."
        ) and registry_entry.unique_id.endswith("-mediaplayer"):
            _LOGGER.debug("Migrating media_player entity unique id")
            return {"new_unique_id": f"{registry_entry.config_entry_id}-projector"}

        if registry_entry.unique_id.startswith(f"{projector.unique_id}-"):
            new_unique_id = registry_entry.unique_id.replace(
                f"{projector.unique_id}-", f"{registry_entry.config_entry_id}-"
            )
            _LOGGER.debug("Migrating entity unique id")
            return {"new_unique_id": new_unique_id}

        # No migration needed
        return None

    await entity_registry.async_migrate_entries(
        hass, entry.entry_id, _async_migrate_entity_entry
    )

    # Open the connection for non-telnet transports. Telnet candidate is already connected.
    if conf_type != CONF_TYPE_TELNET:
        if not await projector.connect(interval=interval):
            raise ConfigEntryNotReady(
                f"Unable to connect to device {projector.unique_id}"
            )

    _LOGGER.info("Device %s is available", projector.unique_id)

    coordinator = BenQProjectorCoordinator(hass, projector, poll_interval=interval)
    coordinator.async_start_state_polling()

    hass.data.setdefault(DOMAIN, {})[entry.entry_id] = coordinator

    await hass.config_entries.async_forward_entry_setups(entry, PLATFORMS)

    entry.async_on_unload(entry.add_update_listener(update_listener))

    async def async_handle_send(call: ServiceCall):
        """Handle the send service call."""
        command: str = call.data.get(CONF_SERVICE_COMMAND)
        action: str = call.data.get(CONF_SERVICE_ACTION)

        response = await coordinator.projector.send_command(command, action, False)

        return {"response": response}

    async def async_handle_send_raw(call: ServiceCall):
        """Handle the send_raw service call."""
        command: str = call.data.get(CONF_SERVICE_COMMAND)

        response = await coordinator.async_send_raw_command(command)

        return {"response": response}

    hass.services.async_register(
        DOMAIN,
        "send",
        async_handle_send,
        schema=SERVICE_SEND_SCHEMA,
        supports_response=SupportsResponse.ONLY,
    )

    hass.services.async_register(
        DOMAIN,
        "send_raw",
        async_handle_send_raw,
        schema=SERVICE_SEND_RAW_SCHEMA,
        supports_response=SupportsResponse.ONLY,
    )

    return True


async def async_unload_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Unload a config entry."""
    coordinator: BenQProjectorCoordinator = hass.data[DOMAIN][entry.entry_id]
    await coordinator.async_disconnect()

    if unload_ok := await hass.config_entries.async_unload_platforms(entry, PLATFORMS):
        hass.data[DOMAIN].pop(entry.entry_id)

    return unload_ok


async def update_listener(hass: HomeAssistant, entry: ConfigEntry) -> None:
    """Handle options update."""
    hass.config_entries.async_schedule_reload(entry.entry_id)
