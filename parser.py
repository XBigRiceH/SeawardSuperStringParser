import logging
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter

from record_types import *

logger = logging.getLogger()
logger.setLevel(logging.INFO)
console_handler = logging.StreamHandler()
formatter = logging.Formatter(
    fmt='%(asctime)s %(levelname)-5s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        logger.info(f"Using .sss file: {file_path}")
    else:
        logger.critical("Usage: python parser.py <sss_file_path>")
        exit()

    f = open(file_path, 'rb')
    test_results: List[TestResult] = []
    machine_info = None

    # parsing
    while True:
        start_byte = f.read(1)
        assert start_byte == b'\x55'

        record_length = f.read(2)
        length_val = struct.unpack('<H', record_length)[0]
        logger.debug(f'Found record, length = {length_val}')

        checksum = f.read(2)
        checksum_val = struct.unpack('<H', checksum)[0]

        empty_byte = f.read(2)
        assert empty_byte == b'\x00\x00'

        # I don't know why but sometimes the length will short a little bit
        record_content = f.read(length_val)
        calculated_checksum = (sum(record_content)) & 0xffff
        if calculated_checksum == checksum_val or calculated_checksum == checksum_val + 1:
            logger.debug(f'Record checksum passed')
        else:
            length_val += 1
            record_content += f.read(1)
            calculated_checksum = (sum(record_content)) & 0xffff
            assert calculated_checksum == checksum_val or calculated_checksum == checksum_val + 1
            logger.debug(f'Record checksum passed')

        record_type = record_content[0]
        if record_type == 0xaa and record_content[1] == 0xff:
            logger.info(f"Parsed {len(test_results)} record, ready to write")
            break

        instance = record_type_class_defs[record_type](record_content[1:])
        if type(instance) is MachineInfo:
            machine_info = instance
        else:
            test_results.append(instance)
        logger.debug(f'Record content = {instance}')

    # writing into Excel file
    logger.info(f'Writing into Excel file...')
    wb = Workbook()
    ws = wb.active
    fill = PatternFill(start_color="F56C6C", end_color="F56C6C", fill_type="solid")

    # Instrument info
    ws.append(
        ["Test Instrument Model", "", "", machine_info.machine_model, "", "", "", "Test Instrument Serial Number", "",
         "", machine_info.machine_serial_number, "", "", ""])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=7)
    ws.merge_cells(start_row=1, start_column=8, end_row=1, end_column=10)
    ws.merge_cells(start_row=1, start_column=11, end_row=1, end_column=14)

    # headers
    headers = [
        "Asset ID", "Site Name", "Location Name", "Test Time",
        "Test Operator", "Overall Result", "Program", "Comments",
        "Next Full Test Date", "Next Formal Visual Test Date", "Test Result"
    ]
    ws.append(headers)

    # sub headers
    sub_headers = ["", "", "", "", "", "", "", "", "", "", "Test Type", "Result", "Unit", "Status"]
    ws.append(sub_headers)

    for col in range(1, 11):
        ws.merge_cells(start_row=2, start_column=col, end_row=3, end_column=col)
    ws.merge_cells(start_row=2, start_column=11, end_row=2, end_column=14)

    current_row = 4
    for record in test_results:
        formatted_results = []
        used_row = 0
        merge_row = []
        failed_row = []
        for visual_test_result in record.visual_test_results:
            formatted_results.append([
                visual_test_result.name,
                visual_test_result.result if visual_test_result.unit else "",
                visual_test_result.unit,
                visual_test_result.flags[0]
            ])
            if not visual_test_result.unit:
                merge_row.append(current_row + used_row)
            if visual_test_result.flags[0] == 'FAIL':
                failed_row.append(current_row + used_row)
            used_row += 1
        for physical_test_result in record.physical_test_results:
            if isinstance(physical_test_result, EarthResistanceTestResult):
                formatted_results.append([
                    "Earth Continuity",
                    physical_test_result.get_value(),
                    physical_test_result.resistance.unit,
                    physical_test_result.get_status()
                ])
                if physical_test_result.get_status() == 'FAIL':
                    failed_row.append(current_row + used_row)
                used_row += 1
            elif isinstance(physical_test_result, IECLeadContinuityTestResult):
                formatted_results.append([
                    "IEC Lead Continuity",
                    physical_test_result.get_value(),
                    physical_test_result.resistance.unit,
                    physical_test_result.get_status()
                ])
                if physical_test_result.get_status() == 'FAIL':
                    failed_row.append(current_row + used_row)
                used_row += 1
            elif isinstance(physical_test_result, PointToPointTestResult):
                formatted_results.append([
                    "Point To Point Resistance",
                    physical_test_result.get_value(),
                    physical_test_result.resistance.unit,
                    physical_test_result.get_status()
                ])
                if physical_test_result.get_status() == 'FAIL':
                    failed_row.append(current_row + used_row)
                used_row += 1
            elif isinstance(physical_test_result, InsulationTestResult):
                formatted_results.append([
                    "Insulation",
                    physical_test_result.get_value(),
                    physical_test_result.resistance.unit,
                    physical_test_result.get_status()
                ])
                if physical_test_result.get_status() == 'FAIL':
                    failed_row.append(current_row + used_row)
                formatted_results.append([
                    "Insulation Voltage",
                    physical_test_result.voltage.value,
                    physical_test_result.voltage.unit,
                    "INFORMATION"
                ])
                used_row += 2
            elif isinstance(physical_test_result, SubstituteLeakageTestResult):
                formatted_results.append([
                    "Substitute Leakage Current",
                    physical_test_result.get_value(),
                    physical_test_result.current.unit,
                    physical_test_result.get_status()
                ])
                if physical_test_result.get_status() == 'FAIL':
                    failed_row.append(current_row + used_row)
                used_row += 1
            elif isinstance(physical_test_result, PolarityTestResult):
                merge_row.append(current_row + used_row)
                formatted_results.append([
                    "IEC Lead Polarity",
                    physical_test_result.get_value(),
                    "",
                    physical_test_result.get_status()
                ])
                if physical_test_result.get_status() == 'FAIL':
                    failed_row.append(current_row + used_row)
                used_row += 1
            elif isinstance(physical_test_result, MainVoltageTestResult):
                formatted_results.append([
                    "Main Voltage",
                    physical_test_result.get_value(),
                    physical_test_result.voltage.unit,
                    physical_test_result.get_status()
                ])
                if physical_test_result.get_status() == 'FAIL':
                    failed_row.append(current_row + used_row)
                used_row += 1
            elif isinstance(physical_test_result, TouchOrLeakageCurrentTestResult):
                formatted_results.append([
                    "Touch Or Leakage Test Load Current",
                    physical_test_result.load_current.value,
                    physical_test_result.load_current.unit,
                    physical_test_result.get_status()
                ])
                formatted_results.append([
                    "Touch Or Leakage Test Leakage Current",
                    physical_test_result.leakage_current.value,
                    physical_test_result.leakage_current.unit,
                    physical_test_result.get_status()
                ])
                if physical_test_result.get_status() == 'FAIL':
                    failed_row.append(current_row + used_row)
                    failed_row.append(current_row + used_row + 1)
                used_row += 2
            elif isinstance(physical_test_result, RCDTestResult):
                formatted_results.append([
                    "RCD Test Current",
                    physical_test_result.test_current.value,
                    physical_test_result.test_current.unit,
                    "INFORMATION"
                ])
                formatted_results.append([
                    "RCD Test Circle Angle",
                    physical_test_result.circle_angle.value,
                    physical_test_result.circle_angle.unit,
                    "INFORMATION"
                ])
                formatted_results.append([
                    "RCD Test Trip time",
                    physical_test_result.get_value(),
                    physical_test_result.trip_time.unit,
                    physical_test_result.get_status()
                ])
                if physical_test_result.get_status() == 'FAIL':
                    failed_row.append(used_row + 2)
                used_row += 3
            elif isinstance(physical_test_result, StringComment):
                merge_row.append(current_row + used_row)
                formatted_results.append([
                    physical_test_result.string_value,
                    "",
                    "",
                    physical_test_result.get_status()
                ])
                used_row += 1

        for i, tr in enumerate(formatted_results):
            row = []
            if i == 0:
                row.extend([
                    record.asset_id,
                    record.site_name,
                    record.location_name,
                    record.test_time,
                    record.test_operator,
                    record.get_status(),
                    record.program,
                    record.comments,
                    record.next_full_test_date,
                    record.next_formal_visual_test_date
                ])
            else:
                row.extend([""] * 10)

            row.extend(tr)
            ws.append(row)

        # merge information cols
        if used_row > 1:
            for col in range(1, 11):
                ws.merge_cells(
                    start_row=current_row,
                    start_column=col,
                    end_row=current_row + used_row - 1,
                    end_column=col
                )
        for i in merge_row:
            ws.merge_cells(
                start_row=i,
                start_column=12,
                end_row=i,
                end_column=13
            )

        # highlight fail cells
        if len(failed_row) > 0:
            ws.cell(row=current_row, column=1).fill = fill
            for row in failed_row:
                for col in range(11, 15):
                    ws.cell(row=row, column=col).fill = fill
        current_row += used_row

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    result_file_path = f"{os.path.splitext(os.path.basename(file_path))[0]}_parsed_{datetime.now().strftime('%y_%m_%d_%H_%M_%S')}.xlsx"
    wb.save(result_file_path)

    logger.info(
        f"All test results have been written to {result_file_path}, total lines = {current_row - 4}, total records = {len(test_results)}")
