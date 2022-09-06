# -*- coding: utf-8 -*-
from typing import Tuple

import pyvisa as visa


class VNA:
    def __init__(self, address) -> None:
        self._rm = visa.ResourceManager()
        self._instr = self._rm.open_resource(address)

    def close(self) -> Tuple[bool, str]:
        """Disconnect the instrument."""
        try:
            self._instr.close()
            self._instr = None
        except Exception as e:
            return (False, str(e))

    def active_channel(self, ch: int) -> Tuple[bool, str]:
        """Specifies selected channel (Ch) as the active channel."""
        self._instr.write(":DISP:WIND{ch}:ACT".format(ch))
        return (True, "OK")

    def active_trace(self, tr: int, ch: int = 1) -> Tuple[bool, str]:
        """Sets/gets the selected trace(Tr) of selected channel(Ch) to the active trace."""
        self._instr.write(":CALC{ch}:PAR{tr}:SEL".format(tr, ch))
        return (True, "OK")

    def trace_number(self, num: int, ch: int = 1) -> Tuple[bool, str]:
        """Sets/gets the number of traces of selected channel(Ch)."""
        self._instr.write(":CALC{ch}:PAR:COUN {num}".format(num, ch))
        return (True, "OK")

    def window_layout(self, layout: str) -> Tuple[bool, str]:
        """Sets the layout of the channel windows on the LCD display."""
        self._instr.write(":DISP:SPL {}".format(layout))
        return (True, "OK")

