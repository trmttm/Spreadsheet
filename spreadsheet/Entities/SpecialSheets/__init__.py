from . import implementation as impl
from .SensitivitySheet import SensitivitySheet
from .VBAcommands import VBAcommands

def wrapper_equal_sign(addresses: tuple) -> tuple:
    return tuple(f'={a}' for a in addresses)
