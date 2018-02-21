import binascii
import struct
import serial
import time


class EdfaOFP2():
    """Commands for OFP2 Module"""

    def __init__(self, serial_address):
        self.address = "COM" + str(serial_address)
        self.open_serial = serial.Serial(self.address, 115200)
        print(self.address)

    def read_reg(self, register_address, length):
        read_num = int(length) + 1
        tx_data = bytes.fromhex(
            '44 A1' + ' ' + register_address + ' ' + length)
        self.open_serial.write(tx_data)
        rx_data = str(binascii.b2a_hex(
            self.open_serial.read(read_num)))[2:-1]
        time.sleep(0.1)
        return rx_data

    def write_reg(self, register_address, value):
        tx_data = bytes.fromhex('44 A0' + ' ' + register_address + ' ' + value)
        self.open_serial.write(tx_data)
        rx_data = str(binascii.b2a_hex(
            self.open_serial.read()))[2:-1]
        time.sleep(0.1)
        return rx_data

    def read_pd(self, pd_id):
        pd_address = {'PDT1': '0C', 'PDT2': '0E', 'PD1': '10', 'PD2': '12',
                      'PD3': '14', 'PD4': '16', 'PD5': '18', 'PD6': '1A',
                      'PD7': '1C', 'PD8': '1E', 'PD9': '20', 'PD10': '22',
                      'PD11': '24', 'PD12': '26', 'PD13': '28', 'PD14': '2A',
                      'PD15': '2C', 'PD16': '2E', 'APC1': '0C', 'APC2': '0E'}
        rx_raw_data = self.read_reg(pd_address[pd_id], '01')
        rx_data = "%.2f" % (0.1 * int(rx_raw_data[:], 16))
        if float(rx_data) >= 3276.8:
            rx_data = "%.2f" % (-0.1 *
                                int(hex(65536 - int(rx_raw_data[:], 16)), 16))
        return float(rx_data)

    def set_edfa_mode(self, edfa_id, value):
        register_address = {'E1': '40', 'E2': '50'}
        if value == 'AGC':
            rx_data = self.write_reg(register_address[edfa_id], '00')
        elif value == 'APC':
            rx_data = self.write_reg(register_address[edfa_id], '01')
        else:
            print('The mode is not support.')
        return rx_data

    def set_edfa_gain(self, edfa_id, value):
        register_address = {'E1': '46', 'E2': '56'}
        value = int(10 * value)
        value = str(binascii.b2a_hex(struct.pack(">h", value)))[2:-1].upper()
        value = value[:2] + ' ' + value[2:]
        rx_data = self.write_reg(register_address[edfa_id], value)
        return rx_data

    def set_edfa_power(self, edfa_id, value):
        register_address = {'E1': '44', 'E2': '54'}
        value = int(10 * value)
        value = str(binascii.b2a_hex(struct.pack(">h", value)))[2:-1].upper()
        value = value[:2] + ' ' + value[2:]
        rx_data = self.write_reg(register_address[edfa_id], value)
        return rx_data
