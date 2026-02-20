import struct
from abc import abstractmethod, ABC
from datetime import datetime
from typing import List

from dateutil.relativedelta import relativedelta


class BufferedRecord:
    def __init__(self, data: bytes):
        self.data = data
        self.idx = 0

    def read(self, length: int, increase_idx: bool = True) -> bytes:
        res = self.data[self.idx: self.idx + length]
        if increase_idx:
            self.idx += length
        return res

    def read_str(self, length: int) -> str:
        return struct.unpack(f'{length}s', self.read(length))[0].replace(b'\x00', b'').decode('utf-8').rstrip()

    def read_float16(self):
        raw = struct.unpack('<H', self.read(2))[0]
        exponent = (raw >> 14) & 0b11
        significand = raw & 0x3FFF
        return round(significand * (0.1 ** exponent), 2)

    def read_uint16(self):
        return struct.unpack('<H', self.read(2))[0]

    def read_uint8(self):
        return struct.unpack('<B', self.read(1))[0]

    def read_flag(self):
        byte_val = self.read(1)[0]
        bits = [(byte_val >> i) & 1 for i in range(7, -1, -1)]
        flag_map = {
            "RESULT_GREATER_THAN": bits[2],
            "RESULT_LESS_THAN": bits[3],
            "FAIL": bits[6],
            "PASS": bits[7],
            "INFO": sum(bits) == 0
        }
        active_flags = [k for k, v in flag_map.items() if v]
        return active_flags if active_flags else ["UNKNOWN"]

    def skip(self, length: int):
        self.idx += length

    def __str__(self):
        filtered = {k: v for k, v in self.__dict__.items() if k not in ('data', 'idx')}
        return f"{self.__class__.__name__}({filtered})"

    def __repr__(self):
        return self.__str__()


class ValueWithUnit:
    def __init__(self, value, unit):
        self.value = value
        self.unit = unit

    def __str__(self):
        return f"{self.__class__.__name__}({self.__dict__})"

    def __repr__(self):
        return self.__str__()


class MachineInfo(BufferedRecord):
    def __init__(self, data: bytes):
        super().__init__(data)

        self.machine_model = self.read_str(20)
        self.machine_serial_number = self.read_str(20)


class TestResult(BufferedRecord):
    def __init__(self, data: bytes):
        super().__init__(data)
        self.flags = []

    def get_status(self):
        if 'FAIL' in self.flags:
            return 'FAIL'
        elif 'PASS' in self.flags:
            return 'PASS'
        elif 'INFO' in self.flags:
            return 'INFO'
        else:
            return 'UNKNOWN'

    def parse_value(self, value):
        if 'RESULT_GREATER_THAN' in self.flags:
            return f"> {value}"
        elif 'RESULT_LESS_THAN' in self.flags:
            return f"< {value}"
        else:
            return f"{value}"


class VisualTestResult(TestResult):
    def __init__(self, data: bytes):
        super().__init__(data)

        self.name = self.read_str(16)
        self.unit = self.read_str(16)
        self.result = self.read_float16()
        self.flags = self.read_flag()


class PhysicalTestResult(TestResult, ABC):
    result_length = 0

    def __init__(self, data: bytes):
        super().__init__(data)

    @abstractmethod
    def get_value(self):
        pass


class EarthResistanceTestResult(PhysicalTestResult):
    result_length = 3

    def __init__(self, data: bytes):
        super().__init__(data)

        self.resistance = ValueWithUnit(self.read_float16(), 'ohm')
        self.flags = self.read_flag()

    def get_value(self):
        return self.parse_value(self.resistance.value)


class IECLeadContinuityTestResult(PhysicalTestResult):
    result_length = 3

    def __init__(self, data: bytes):
        super().__init__(data)

        self.resistance = ValueWithUnit(self.read_float16(), 'ohm')
        self.flags = self.read_flag()

    def get_value(self):
        return self.parse_value(self.resistance.value)


class PointToPointTestResult(PhysicalTestResult):
    result_length = 3

    def __init__(self, data: bytes):
        super().__init__(data)

        self.resistance = ValueWithUnit(self.read_float16(), 'ohm')
        self.flags = self.read_flag()

    def get_value(self):
        return self.parse_value(self.resistance.value)


class InsulationTestResult(PhysicalTestResult):
    result_length = 5

    def __init__(self, data: bytes):
        super().__init__(data)

        self.voltage = ValueWithUnit(self.read_float16(), 'v')
        self.resistance = ValueWithUnit(self.read_float16(), 'mohm')
        self.flags = self.read_flag()

    def get_value(self):
        return self.parse_value(self.resistance.value)


class SubstituteLeakageTestResult(PhysicalTestResult):
    result_length = 3

    def __init__(self, data: bytes):
        super().__init__(data)

        self.current = ValueWithUnit(self.read_float16(), 'ma')
        self.flags = self.read_flag()

    def get_value(self):
        return self.parse_value(self.current.value)


class PolarityTestResult(PhysicalTestResult):
    result_length = 1

    def __init__(self, data: bytes):
        super().__init__(data)

        self.flags = self.read_flag()

    def get_value(self):
        if 'FAIL' in self.flags:
            return 'Live / Neutral Reversed'
        return ''


class MainVoltageTestResult(PhysicalTestResult):
    result_length = 3

    def __init__(self, data: bytes):
        super().__init__(data)

        self.voltage = ValueWithUnit(self.read_float16(), 'v')
        self.flags = self.read_flag()

    def get_value(self):
        return self.parse_value(self.voltage.value)


class TouchOrLeakageCurrentTestResult(PhysicalTestResult):
    result_length = 7

    def __init__(self, data: bytes):
        super().__init__(data)

        self.load_current = ValueWithUnit(self.read_float16(), 'ma')
        self.skip(2)
        self.leakage_current = ValueWithUnit(self.read_float16(), 'ma')
        self.flags = self.read_flag()

    def get_value(self):
        return ''


class RCDTestResult(PhysicalTestResult):
    result_length = 7

    def __init__(self, data: bytes):
        super().__init__(data)

        self.test_current = ValueWithUnit(self.read_float16(), 'ma')
        self.circle_angle = ValueWithUnit(self.read_float16(), 'deg')
        self.trip_time = ValueWithUnit(self.read_float16(), 'ms')
        self.flags = self.read_flag()

    def get_value(self):
        return self.parse_value(self.trip_time.value)


class StringComment(PhysicalTestResult):
    result_length = 87

    def __init__(self, data: bytes):
        super().__init__(data)

        self.string_value = self.read_str(86)
        self.flags = self.read_flag()

    def get_value(self):
        return ''


physical_test_type_class_defs = {
    0x11: EarthResistanceTestResult,
    0x16: IECLeadContinuityTestResult,
    0x18: PointToPointTestResult,
    0x20: InsulationTestResult,
    0x83: SubstituteLeakageTestResult,
    0x91: PolarityTestResult,
    0x92: MainVoltageTestResult,
    0x96: TouchOrLeakageCurrentTestResult,
    0x9A: RCDTestResult,
    0xfc: StringComment
}


class TestResult(BufferedRecord):
    def __init__(self, data: bytes):
        super().__init__(data)
        self.flags = self.read_flag()
        self.asset_id = self.read_str(16)
        self.skip(64)
        self.site_name = self.read_str(16)
        self.location_name = self.read_str(16)
        self.test_time = datetime(
            hour=self.read_uint8(),
            minute=self.read_uint8(),
            second=self.read_uint8(),
            day=self.read_uint8(),
            month=self.read_uint8(),
            year=self.read_uint16()
        )
        self.test_operator = self.read_str(16)
        self.comments = self.read_str(128)
        self.skip(1)
        self.next_full_test_date = self.test_time + relativedelta(months=self.read_uint8())
        self.program = self.read_str(30)
        self.next_formal_visual_test_date = self.test_time + relativedelta(months=self.read_uint8())
        self.skip(15)

        while self.read(1)[0] != 0xfe:
            pass

        self.visual_test_results: List[VisualTestResult] = []
        self.physical_test_results: List[PhysicalTestResult] = []

        while self.idx < len(self.data) - 2:
            test_type = self.read(1)[0]
            if test_type == 0xfd:
                self.visual_test_results.append(VisualTestResult(self.read(35)))
            elif test_type in physical_test_type_class_defs.keys():
                self.physical_test_results.append(physical_test_type_class_defs[test_type](
                    self.read(physical_test_type_class_defs[test_type].result_length)))
            else:
                print('Unknown test type:', self.read(30))
                return

    def get_status(self):
        return self.flags[0]


record_type_defs = {
    b'\x55': 'Machine Info',
    b'\x01': 'Test Result',
}

record_type_class_defs = {
    0x55: MachineInfo,
    0x01: TestResult
}
