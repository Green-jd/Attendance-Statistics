import json

import xlrd

from Record import Record, Person, Attendance
from Times import Times


def get_time(time_cell):
    duration = time_cell.value
    times = duration.split(' ~ ')
    start_time = times[0].split('/')
    end_time = times[1][0:5].split('/')

    year = int(start_time[0])
    start_month = int(start_time[1])
    start_day = int(start_time[2])

    end_month = int(end_time[0])
    end_day = int(end_time[1])

    return Times(year, start_month, start_day, end_month, end_day)


def get_member_record(record_time, member_info_cell, member_records_cell):
    user_number = member_info_cell[2].value
    user_name = member_info_cell[10].value

    count = record_time.end_day - record_time.start_day + 1

    records = []
    for index in range(0, count):

        record_source = member_records_cell[index].value
        check_in = ''
        check_out = ''

        if len(record_source) > 0:
            detail = record_source.split('\n')
            # print(f'value = {detail}')
            if len(detail) == 2:
                # 只有一条记录的时候需要判断是签到还是签退
                if int(detail[0].split(':')[0]) < 12:
                    #  早于12点认为是签到
                    check_in = detail[0]
                else:
                    check_out = detail[0]
            else:
                check_in = detail[0]
                check_out = detail[1]

        record_detail = Record(str(record_time.start_month) + '/' + str(index + 1), check_in, check_out)
        records.append(record_detail)
        # print(
        #     f'name = {user_name}, number = {user_number}, day = {record_time.start_month}/{index + 1}, check_in = {check_in}, check_out = {check_out}')

    return Person(user_name, user_number, records)


def calculate_records(filename):
    wb = xlrd.open_workbook(filename)
    record_sheet = wb.sheet_by_name('刷卡记录')

    record_time = get_time(record_sheet.cell(2, 2))
    print(
        f'{record_time.year}/{record_time.start_month}/{record_time.start_day} - {record_time.year}/{record_time.end_month}/{record_time.end_day}')

    member = 4
    records = []
    while (member - 4) * 2 + 5 < record_sheet.nrows:
        member_record = get_member_record(record_time,
                                          record_sheet.row((member - 4) * 2 + 4),
                                          record_sheet.row((member - 4) * 2 + 5))
        member += 1
        records.append(member_record)
        print(member_record.name)
        print(json.dumps(member_record, default=lambda obj: obj.__dict__))

    return Attendance(record_time, records)
