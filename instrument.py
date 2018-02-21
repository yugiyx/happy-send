import time
import visa
import xlrd
import xlwt
from xlutils.copy import copy

rm = visa.ResourceManager()


class Att8156A():
    """This is a Agilent 8156A Attenuator.8156A have only 1 slot."""

    def __init__(self, gpib_address):
        self.address = "GPIB0::" + str(gpib_address) + "::INSTR"
        self.open_gpib = rm.open_resource(self.address)

    def get_att_value(self):
        tx_data = ':INP' + ':ATT?'
        rx_data = float("%.3f" % float(self.open_gpib.query(tx_data)))
        return rx_data

    def set_att_value(self, set_value):
        tx_data = ':INP' + ':ATT' + ' ' + str(set_value)
        self.open_gpib.write(tx_data)
        return 'Done'

    def set_att_wavlength(self, set_value):
        tx_data = ':INP' + ':WAV' + ' ' + str(set_value)
        self.open_gpib.write(tx_data)
        return 'Done'

    def control_att_shutter(self, set_value):
        tx_data = ':OUTP' + ':STAT' + ' ' + str(set_value)
        self.open_gpib.write(tx_data)
        return 'Done'


class Pm8163A():
    """This is a Agilent 8163A Power Meter.8163A have 2 slot."""

    def __init__(self, gpib_address):
        self.address = "GPIB0::" + str(gpib_address) + "::INSTR"
        self.open_gpib = rm.open_resource(self.address)

    def get_pm_value(self, slot='1', channel='1'):
        tx_data = ':READ' + str(slot) + ':CHAN' + str(channel) + ':POW?'
        rx_data = float("%.3f" % float(self.open_gpib.query(tx_data)))
        return rx_data

    def set_pm_wavlength(self, set_value, slot='1', channel='1'):
        tx_data = ':SENS' + str(slot) + ':CHAN' + \
            str(channel) + ':POW:WAV' + ' ' + str(set_value)
        self.open_gpib.write(tx_data)
        return 'Done'


class AttPmN7752A():
    """This is a Agilent N7752A Attenuators and Power Meter.
    Slot 1 to 2 is Attenuator 1.
    Slot 3 to 4 is Attenuator 2.
    Slot 5 and 6 are Power Meter 1 and 2."""

    def __init__(self, gpib_address):
        self.address = "GPIB0::" + str(gpib_address) + "::INSTR"
        self.open_gpib = rm.open_resource(self.address)

    def get_att_value(self, slot='1', channel='1'):
        tx_data = ':INP' + str(slot) + ':CHAN' + str(channel) + ':ATT?'
        rx_data = float("%.3f" % float(self.open_gpib.query(tx_data)))
        return rx_data

    def set_att_value(self, set_value, slot='1', channel='1'):
        tx_data = ':INP' + str(slot) + ':CHAN' + \
            str(channel) + ':ATT' + ' ' + str(set_value)
        self.open_gpib.write(tx_data)
        return 'Done'

    def set_att_wavlength(self, set_value, slot='1', channel='1'):
        tx_data = ':INP' + str(slot) + ':CHAN' + \
            str(channel) + ':WAV' + ' ' + str(set_value)
        self.open_gpib.write(tx_data)
        return 'Done'

    def control_att_shutter(self, set_value, slot='1', channel='1'):
        tx_data = ':OUTP' + str(slot) + ':STAT' + ' ' + str(set_value)
        self.open_gpib.write(tx_data)
        return 'Done'

    def set_apc_value(self, set_value, slot='1', channel='1'):
        tx_data = 'OUTP' + str(slot) + ':CHAN' + \
            str(channel) + ':POW' + ' ' + str(set_value)
        self.open_gpib.write(tx_data)
        return 'Done'

    def enable_apc_mode(self, set_value, slot='1', channel='1'):
        tx_data = 'OUTP' + str(slot) + ':CHAN' + str(channel) + \
            ':POW:CONTR' + ' ' + str(set_value)
        print(tx_data)
        self.open_gpib.write(tx_data)
        return 'Done'

    def get_pm_value(self, slot='1', channel='1'):
        tx_data = 'READ' + str(slot) + ':CHAN' + str(channel) + ':POW?'
        rx_data = float("%.3f" % float(self.open_gpib.query(tx_data)))
        return rx_data

    def set_pm_wavlength(self, set_value, slot='1', channel='1'):
        tx_data = ':SENS' + str(slot) + ':CHAN' + \
            str(channel) + ':POW:WAV' + ' ' + str(set_value)
        self.open_gpib.write(tx_data)
        return 'Done'


class OsaAQ6370():
    """This is a Yokogawa AQ6370 Optical Spectrum Analyzer."""

    def __init__(self, gpib_address):
        self.address = "GPIB0::" + str(gpib_address) + "::INSTR"
        self.open_gpib = rm.open_resource(self.address)

    def sweep_trace(self, trace, trace_offset):
        tx_data = ':TRAC:ACT TR' + str(trace)
        self.open_gpib.write(tx_data)
        tx_data = ':TRAC:ATTR:TR' + str(trace) + ' ' + 'WRIT'
        self.open_gpib.write(tx_data)
        tx_data = ':INIT:SMOD SING'
        self.open_gpib.write(tx_data)
        tx_data = ':INIT'
        self.open_gpib.write(tx_data)
        time.sleep(3)
        tx_data = ':TRAC:ATTR:TR' + str(trace) + ' ' + 'FIX'
        self.open_gpib.write(tx_data)
        tx_data = ':CALC:CAT POW'
        self.open_gpib.write(tx_data)
        tx_data = ':CALC:PAR:POW:OFFS ' + str(trace_offset)
        self.open_gpib.write(tx_data)
        tx_data = ':CALC'
        self.open_gpib.write(tx_data)
        raw_data = float(
            '%.3f' % (float(self.open_gpib.query(':CALC:DATA?'))))
        return raw_data

    def analysis_nf(self):
        tx_data = ':CALC:CAT NF'
        self.open_gpib.write(tx_data)
        tx_data = ':CALC'
        self.open_gpib.write(tx_data)
        raw_data = self.open_gpib.query(':CALC:DATA?')
        return raw_data


class SwZ4400():
    """This is a Z4400 Switch.
        Use function 'setswch' to set position:
    """

    def __init__(self, serial_address):
        self.address = "ASRL" + str(serial_address) + "::INSTR"
        self.open_serial = rm.open_resource(self.address, baud_rate=115200)

    def set_sw(self, channel, value):
        tx_data = 'setswch' + ' ' + str(channel) + ' ' + str(value)
        self.open_serial.write(tx_data)
        time.sleep(0.3)
        return 'Done'


class SwZHDIY():
    """This is a ZHDIY Switch.
        Use function 'set allsw' to set position:
            'MLS_PM'        'TLS_PM'
            'MLS_OSA'       'TLS_OSA'
            'MLS_EDFA_PM'   'TLS_EDFA_PM'
            'MLS_EDFA_OSA'  'TLS_EDFA_OSA'
    """

    def __init__(self, serial_address):
        self.address = "ASRL" + str(serial_address) + "::INSTR"
        self.open_serial = rm.open_resource(self.address, baud_rate=115200)

    def set_sw(self, postion):
        set_value = {'MLS_PM': 'BAABBAAA', 'TLS_PM': 'AAABBAAA',
                     'MLS_OSA': 'BAABBBAA', 'TLS_OSA': 'AAABBBAA',
                     'MLS_EDFA_PM': 'BBAAAAAA', 'TLS_EDFA_PM': 'ABAAAAAA',
                     'MLS_EDFA_OSA': 'BBAAABAA', 'TLS_EDFA_OSA': 'ABAAABAA', }
        tx_data = 'set allsw' + ' ' + set_value[postion]
        self.open_serial.write(tx_data)
        time.sleep(0.3)
        return 'Done'


class DataProcess():
    """docstring for Data_process"""

    def __init__(self):
        pass

    def open_config(self, xls_name, test_item):
        set_values = []
        book = xlrd.open_workbook(xls_name)
        sheet = book.sheet_by_name(test_item)
        i = 0
        while i < sheet.ncols:
            set_values.append(sheet.col_values(i))
            set_values[i].pop(0)
            i += 1
        return set_values

    def save_pd_data(self, xls_name, test_item, test_data):
        book = xlrd.open_workbook(xls_name)
        wbook = copy(book)
        wsheet = wbook.add_sheet(test_item, cell_overwrite_ok=True)
        nwrows = 0
        nwcols = 0
        i = 0
        for x in test_data[0]:
            wsheet.write(nwrows, nwcols, x)
            for y in range(len(test_data))[1:]:
                wsheet.write(nwrows, nwcols + y, test_data[y][i])
            nwrows += 1
            i += 1
        wbook.save(xls_name)

    def save_nf_data(self, xls_name, a_data, b_data, nf_data):
        wbook = xlwt.Workbook()
        sht = wbook.add_sheet('test_result', cell_overwrite_ok=True)
        i = 0
        m = 5
        n = 21
        nwrows = 1
        nwcols = 1
        while i < 47:
            j = 0
            osa_data_buffer = []
            while j < 7:
                if j == 0 or j == 4:
                    osa_data_buffer.append(
                        '%.3f' % ((float(nf_data[m:n])) * 10**9))
                else:
                    osa_data_buffer.append('%.3f' % (float(nf_data[m:n])))
                m += 17
                n += 17
                j += 1
            i += 1
            sht.write(nwrows, nwcols - 1, i)
            sht.write(nwrows, nwcols + 0, float(osa_data_buffer[0]))
            sht.write(nwrows, nwcols + 1, float(osa_data_buffer[1]))
            sht.write(nwrows, nwcols + 2, float(osa_data_buffer[2]))
            sht.write(nwrows, nwcols + 3, float(osa_data_buffer[3]))
            sht.write(nwrows, nwcols + 4, float(osa_data_buffer[4]))
            sht.write(nwrows, nwcols + 5, float(osa_data_buffer[5]))
            sht.write(nwrows, nwcols + 6, float(osa_data_buffer[6]))
            nwrows += 1
        sht.write(0, 0, 'ch num')
        sht.write(0, 1, 'center wl')
        sht.write(0, 2, 'input_power')
        sht.write(0, 3, 'output_power')
        sht.write(0, 4, 'ase')
        sht.write(0, 5, 'resoln')
        sht.write(0, 6, 'gain')
        sht.write(0, 7, 'nf')

        sht.write(nwrows, 0, 'total_input')
        sht.write(nwrows, 1, a_data)
        sht.write(nwrows + 1, 0, 'total_total_output')
        sht.write(nwrows + 1, 1, b_data)

        wbook.save(xls_name)
