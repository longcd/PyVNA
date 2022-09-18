# -*- coding: utf-8 -*-
from typing import Union

import pyvisa as visa


class VNA:
    def __init__(self, address) -> None:
        self._rm = visa.ResourceManager()
        self._instr = self._rm.open_resource(address)

    def close(self) -> None:
        """Disconnect the instrument."""
        self._instr.close()
        self._instr = None

    ####################
    # ACTIVE CH/TRACE
    ####################
    def active_channel(self, ch: int = 1) -> None:
        """Specifies selected channel (Ch) as the active channel."""
        self._instr.write(f":DISP:WIND{ch}:ACT")

    def active_trace(self, tr: int, ch: int = 1) -> None:
        """Sets/gets the selected trace(Tr) of selected channel(Ch) to the active trace."""
        self._instr.write(f":CALC{ch}:PAR{tr}:SEL")

    def trace_number(self, num: int, ch: int = 1) -> None:
        """Sets/gets the number of traces of selected channel(Ch)."""
        self._instr.write(f":CALC{ch}:PAR:COUN {num}")

    ####################
    # Frequency
    ####################
    def freq_start(self, freq: Union[float, str], ch: int = 1) -> str:
        """Sets/gets the start value of the sweep range of selected channel(Ch).

        Parameters
        ----------
        freq: float or str
            Get the start value if freq is '?'
            Set the start value if freq is a float number, the unit is Hz(hertz).
        """
        if isinstance(freq, str) and freq == "?":
            return self._instr.query(f":SENS{ch}:FREQ:STAR?")
        else:
            return self._instr.write(f":SENS{ch}:FREQ:STAR {freq}")

    def freq_stop(self, freq: Union[float, str], ch: int = 1) -> str:
        """Sets/gets the stop value of the sweep range of selected channel(Ch).

        Parameters
        ----------
        freq: float or str
            Get the stop value if freq is '?'
            Set the stop value if freq is a float number, the unit is Hz(hertz).
        """
        if isinstance(freq, str) and freq == "?":
            return self._instr.query(f":SENS{ch}:FREQ:STAR?")
        else:
            return self._instr.write(f":SENS{ch}:FREQ:STOP {freq}")

    def freq_center(self, freq: Union[float, str], ch: int = 1) -> str:
        """Sets/gets the center value of the sweep range of selected channel(Ch).

        Parameters
        ----------
        freq: float or str
            Get the center value if freq is '?'
            Set the center value if freq is a float number, the unit is Hz(hertz).
        """
        if isinstance(freq, str) and freq == "?":
            return self._instr.query(f":SENS{ch}:FREQ:CENT?")
        else:
            return self._instr.write(f":SENS{ch}:FREQ:CENT {freq}")

    def freq_span(self, freq: Union[float, str], ch: int = 1) -> str:
        """Sets/gets the span value of the sweep range of selected channel(Ch).

        Parameters
        ----------
        freq: float or str
            Get the span value if freq is '?'
            Set the span value if freq is a float number, the unit is Hz(hertz).
        """
        if isinstance(freq, str) and freq == "?":
            return self._instr.query(f":SENS1:FREQ:SPAN?")
        else:
            return self._instr.write(f":SENS{ch}:FREQ:SPAN {freq}")

    ####################
    # Sweep Setup
    ####################
    def power(self, power: Union[float, str], ch: int = 1) -> str:
        """Sets/gets the power level of the selected channel(Ch).

        Parameters
        ----------
        power : float or str
            Get the power level if power is '?'
            Set the power level if power is a float number, the unit is dBm.
        """
        if isinstance(power, str) and power == "?":
            return self._instr.query(f":SOUR{ch}:POW?")
        else:
            return self._instr.write(f":SOUR{ch}:POW {power}")

    def points(self, point: Union[int, str], ch: int = 1) -> str:
        """Sets/gets the number of measurement points of selected channel(Ch).

        Parameters
        ----------
        point : int or str
            Get the points if point is '?'
            Set the points if point is a int number.
        """
        if isinstance(point, str) and point == "?":
            return self._instr.query(f":SENS{ch}:SWE:POIN?")
        else:
            return self._instr.write(f":SENS{ch}:SWE:POIN {point}")

    def sweep_type(self, typ: str, ch: int = 1) -> str:
        """Sets/gets the sweep type of selected channel(Ch).

        Parameters
        ----------
        typ: str
            See: https://rfmw.em.keysight.com/wireless/helpfiles/e5071c/programming/command_reference/sense/scpi_sense_ch_sweep_type.htm
            Select from either of the following:
            - "LINear": Sets the sweep type to the linear sweep.
            - "LOGarithmic": Sets the sweep type to the log sweep.
            - "SEGMent": Sets the sweep type to the segment sweep.
            - "POWer": Sets the sweep type to the power sweep.
        """
        if typ == "?":
            return self._instr.query(f":SENS{ch}:SWE:TYPE?")
        else:
            return self._instr.write(f":SENS{ch}:SWE:TYPE {typ}")

    def segm_data(self, data: str, ch=1) -> str:
        """Creates the segment sweep table of selected channel(Ch).

        Parameters
        ----------
        data: str
            See: https://rfmw.em.keysight.com/wireless/helpfiles/e5071c/programming/command_reference/sense/scpi_sense_ch_segment_data.htm
        """
        if data == "?":
            return self._instr.query(f":SENS{ch}:SEGM:DATA?")
        else:
            return self._instr.write(f":SENS{ch}:SEGM:DATA {data}")

    ####################
    # RESPONSE
    ####################
    def bandwidth(self, bw: Union[float, str], ch: int = 1) -> str:
        """Sets/gets the IF bandwidth of selected channel(Ch).

        Parameters
        ----------
        bw: float or str
            Get the IF bandwidth if bw is '?'
            Set the IF bandwidth if bw is a float number.
        """
        if isinstance(bw, str) and bw == "?":
            return self._instr.query(f":SENS{ch}:BAND?")
        else:
            return self._instr.write(f":SENS{ch}:BAND {bw}")

    def trace_num(self, num: Union[int, str], ch: int = 1) -> str:
        """Sets/gets the number of traces of selected channel(Ch).

        Parameters
        ----------
        num: int or str
            Get the number of traces if num is '?'
            Set the number of traces if num is a int number.
        """
        if isinstance(num, str) and num == "?":
            return self._instr.query(f":CALC{ch}:PAR:COUN?")
        else:
            return self._instr.write(f":CALC{ch}:PAR:COUN {num}")

    def window_layout(self, layout: str, ch: int = 1) -> str:
        """Sets/gets the graph layout of selected channel(Ch).

        Parameters
        ----------
        layout: str
            See: https://rfmw.em.keysight.com/wireless/helpfiles/e5071c/programming/command_reference/display/scpi_display_window_ch_split.htm
        """
        if layout == "?":
            return self._instr.query(":DISP:WIND{ch}:SPL?")
        else:
            return self._instr.write(f":DISP:WIND{ch}:SPL {layout}")

    def parameter(self, para: str, tr: int, ch: int = 1) -> str:
        """Sets/gets the measurement parameter of the selected trace(Tr), for the selected channel(Ch).

        Parameters
        ----------
        parap: str
            See: https://rfmw.em.keysight.com/wireless/helpfiles/e5071c/programming/command_reference/calculate/scpi_calculate_ch_parameter_tr_define.htm
        """
        if para == "?":
            return self._instr.query(f":CALC{ch}:PAR{tr}:DEF?")
        else:
            return self._instr.write(f":CALC{ch}:PAR{tr}:DEF {para}")

    def format(self, form: str, tr: int, ch: int = 1) -> str:
        """Sets/gets the data format of the active trace of selected channel(Ch).

        Parameters
        ----------
        form: str
            See: https://rfmw.em.keysight.com/wireless/helpfiles/e5071c/programming/command_reference/calculate/scpi_calculate_ch_selected_format.htm
        """
        if form == "?":
            return self._instr.query(f":CALC{ch}:TRAC{tr}:FORM?")
        else:
            return self._instr.write(f":CALC{ch}:TRAC{tr}:FORM {form}")

    ####################
    # Calibration
    ####################
    def cal_kit(self, kit: Union[int, str], ch: int = 1) -> str:
        """Sets/gets the calibration kit of selected channel(Ch).

        Parameters
        ----------
        kit: int or str
            Gets the calibration kit if kit is '?'
            Sets the calibration kit if kit is a integer.
        """
        if kit == "?":
            return self._instr.query(f":SENS{ch}:CORR:COLL:CKIT?")
        else:
            return self._instr.write(f":SENS{ch}:CORR:COLL:CKIT {kit}")

    def cal_meth(self, ports: str, ch: int = 1) -> str:
        """Sets the calibration type to the full 2-port calibration between the specified 2 ports, for the selected channel(Ch).

        Parameters
        ----------
        ports: str
            Specifies the port for full 2-port calibration. eg. 1,2
        """
        return self._instr.write(f":SENS{ch}:CORR:COLL:METH:SOLT2 {ports}")

    def cal_open(self, port: int, ch: int = 1) -> None:
        """Measures the calibration data of the open standard for the specified port, for the selected channel(Ch).

        Parameters
        ----------
        port: int
            the specified port
        """
        self._instr.write(f":SENS{ch}:CORR:COLL:OPEN {port}")
        self._instr.query("*OPC?")

    def cal_shor(self, port: int, ch: int = 1) -> None:
        """Measures the calibration data of the short standard for the specified port, for the selected channel(Ch).

        Parameters
        ----------
        port: int
            the specified port
        """
        self._instr.write(f":SENS{ch}:CORR:COLL:SHOR {port}")
        self._instr.query("*OPC?")

    def cal_load(self, port: int, ch: int = 1) -> None:
        """Measures the calibration data of the load standard for the specified port, for the selected channel(Ch).

        Parameters
        ----------
        port: int
            the specified port
        """
        self._instr.write(f":SENS{ch}:CORR:COLL:LOAD {port}")
        self._instr.query("*OPC?")

    def cal(self, typ: str, port: int, ch: int = 1) -> None:
        self._instr.write(f":SENS{ch}:CORR:COLL:{typ} {port}")
        self._instr.query("*OPC?")

    def cal_thru(self, port1: int, port2: int, ch: int = 1) -> None:
        """Measures the calibration data of the Thru standard from the specified stimulus port to the specified response port, for the selected channel(Ch)."""
        self._instr.write(f":SENS{ch}:CORR:COLL:THRU {port1},{port2}")
        self._instr.write(f":SENS{ch}:CORR:COLL:THRU {port2},{port1}")
        self._instr.query("*OPC?")

    def cal_done(self, ch: int = 1) -> None:
        self._instr.write(f":SENS{ch}:CORR:COLL:SAVE")

    ####################
    # Save/Recall
    ####################
    def save_type(self, stype: str) -> str:
        """
        See: https://rfmw.em.keysight.com/wireless/helpfiles/e5071c/programming/command_reference/memory/scpi_mmemory_store_stype.htm
        """
        if stype == "?":
            return self._instr.query(":MMEM:STOR:STYP?")
        else:
            return self._instr.write(f":MMEM:STOR:STYP {stype}")

    def mdir(self, folder: str) -> None:
        self._instr.write(f":MMEM:MDIR '{folder}'")

    def stor_stat(self, file: str) -> None:
        """Saves the instrument state into a file(file with the .sta extension).

        Parameters
        ----------
        file: str
            File name to save the instrument state (extension ".sta")
            eg. 'D:/FILTERS/test.sta'
        """
        self._instr.write(f":MMEM:STOR '{file}'")

    def load_stat(self, file: str) -> None:
        """This command recalls the specified instrument state file.

        Parameters
        ----------
        file: str
            File name of instrument state (extension ".sta")
            eg. 'D:/FILTERS/test.sta'
        """
        self._instr.write(f":MMEM:LOAD '{file}'")

    ####################
    # CALCULATE
    ####################
    def calc_sdat(self, tr: int, ch: int = 1) -> str:
        """Gets the corrected data array, for the active trace of selected channel(Ch)."""
        return self._instr.query(f":CALC{ch}:TRAC{tr}:DATA:SDAT?")

    def calc_fdat(self, tr: int, ch: int = 1) -> str:
        """Gets the formatted data array, for the active trace of selected channel(Ch)."""
        return self._instr.query(f":CALC{ch}:TRAC{tr}:DATA:FDAT?")

    def calc_mfd(self, trs: str, ch: int = 1) -> str:
        """Gets the formatted data array of multiple traces of the selected channel(Ch)."""
        return self._instr.query(f":CALC{ch}:TRAC:DATA:MFD? '{trs}'")

    def init_cont(self, status: str, ch: int = 1) -> None:
        """Turns ON/OFF the continuous initiation mode of selected channel (Ch) in the trigger system."""
        self._instr.write(f":INIT{ch}:CONT {status}")

    def trigger(self, ch: int = 1) -> None:
        self._instr.write(f":INIT{ch}")

    # Limit Test
    def limit_display(self, state: str, tr: int, ch: int = 1) -> str:
        """Turns ON/OFF the limit line display, for the active trace of selected channel(Ch).

        Parameters
        ----------
        state: str
            ON/OFF
        """
        if state == "?":
            return self._instr.query(f":CALC{ch}:TRAC{tr}:LIM:DISP?")
        else:
            return self._instr.write(f":CALC{ch}:TRAC{tr}:LIM:DISP {state}")

    def limit_data(self, data: str, tr: int, ch: int = 1) -> str:
        """Sets/gets the limit table for the limit test, for the active trace of selected channel(Ch).

        See: https://rfmw.em.keysight.com/wireless/helpfiles/e5071c/programming/command_reference/calculate/scpi_calculate_ch_selected_limit_data.htm
        """
        if data == "?":
            return self._instr.query(f":CALC{ch}:TRAC{tr}:LIM:DATA?")
        else:
            return self._instr.write(f":CALC{ch}:TRAC{tr}:LIM:DATA {data}")

    # IEEE4882
    def preset(self) -> None:
        self._instr.write(":SYST:PRES")
