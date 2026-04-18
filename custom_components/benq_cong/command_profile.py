"""Official RS232 command capability profile for BenQ HT3550/W2700 class models."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class CommandCapability:
    """Describe whether a command supports read and/or write operations."""

    read: bool
    write: bool


OFFICIAL_COMMAND_CAPABILITIES: dict[str, CommandCapability] = {
    # Core power and media controls.
    "pow": CommandCapability(read=True, write=True),
    "sour": CommandCapability(read=True, write=True),
    "mute": CommandCapability(read=True, write=True),
    "vol": CommandCapability(read=True, write=True),
    "modelname": CommandCapability(read=True, write=False),
    "ltim": CommandCapability(read=True, write=False),
    "ltim2": CommandCapability(read=False, write=False),
    # Picture and display settings.
    "appmod": CommandCapability(read=True, write=True),
    "con": CommandCapability(read=True, write=True),
    "bri": CommandCapability(read=True, write=True),
    "color": CommandCapability(read=True, write=True),
    "sharp": CommandCapability(read=True, write=True),
    "ct": CommandCapability(read=True, write=True),
    "asp": CommandCapability(read=True, write=True),
    "bc": CommandCapability(read=True, write=True),
    "lampm": CommandCapability(read=True, write=True),
    "3d": CommandCapability(read=True, write=True),
    "pp": CommandCapability(read=True, write=True),
    "blank": CommandCapability(read=True, write=True),
    "freeze": CommandCapability(read=True, write=True),
    # Installation/system controls.
    "directpower": CommandCapability(read=True, write=True),
    "highaltitude": CommandCapability(read=True, write=True),
    # Officially unsupported / not available for this class.
    "audiosour": CommandCapability(read=False, write=False),
    "micvol": CommandCapability(read=False, write=False),
    "autopower": CommandCapability(read=False, write=False),
    "standbynet": CommandCapability(read=False, write=False),
    "standbymic": CommandCapability(read=False, write=False),
    "standbymnt": CommandCapability(read=False, write=False),
    "amxdd": CommandCapability(read=False, write=False),
    "macaddr": CommandCapability(read=False, write=False),
    "cgamut": CommandCapability(read=False, write=False),
    "qas": CommandCapability(read=False, write=False),
    "ins": CommandCapability(read=False, write=False),
    "lpsaver": CommandCapability(read=False, write=False),
    "prjlogincode": CommandCapability(read=False, write=False),
    "broadcasting": CommandCapability(read=False, write=False),
    "menuposition": CommandCapability(read=False, write=False),
    "led": CommandCapability(read=False, write=False),
    "keyst": CommandCapability(read=False, write=False),
    "hkeystone": CommandCapability(read=False, write=False),
    "vkeystone": CommandCapability(read=False, write=False),
    "rgain": CommandCapability(read=False, write=False),
    "ggain": CommandCapability(read=False, write=False),
    "bgain": CommandCapability(read=False, write=False),
    "roffset": CommandCapability(read=False, write=False),
    "goffset": CommandCapability(read=False, write=False),
    "boffset": CommandCapability(read=False, write=False),
    "hdrbri": CommandCapability(read=False, write=False),
}


OFFICIAL_UNSUPPORTED_COMMANDS: set[str] = {
    command
    for command, capability in OFFICIAL_COMMAND_CAPABILITIES.items()
    if not capability.read and not capability.write
}


def supports_command_by_profile(command: str, operation: str = "any") -> bool:
    """Return command support based on official capability table.

    operation: "any" | "read" | "write"
    Unknown commands are treated as supported to preserve backward compatibility.
    """

    capability = OFFICIAL_COMMAND_CAPABILITIES.get(command.lower())
    if capability is None:
        return True

    if operation == "read":
        return capability.read
    if operation == "write":
        return capability.write

    return capability.read or capability.write
