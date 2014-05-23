#*-* coding: utf-8 *-*
from xlwt import Workbook
import os
mark = '\t'


def time_second(time):
    hh = int(time.split(':')[0])
    mm = int(time.split(':')[1])
    ss = int(time.split(':')[2])

    return hh*3600+mm*60+ss


def calc_work_time(time_end, time_start):
    if time_end != ' ':
        if (time_second(time_end)-time_second(time_start)) >= 4*3600:
            return 1
        else:
            return 0.5
    return 0


def main():
    file_name2 = 'report_jicheng.xls'

    sheet_list_merge_file = open('sheetlist_jicheng.txt', 'r')
    sheet_list_merge = sheet_list_merge_file.readline().split('\t')
    sheet_list_merge_file.close()
    number2 = len(sheet_list_merge)

    wb2 = Workbook()
    if os.path.isfile(file_name2):
        print 'This "gps_jicheng.xls" is already, will be deleted'
        os.remove(file_name2)
        print 'The "gps_jicheng.xls" has been deleted'
    print 'write data to excel'

    for index in range(number2):
        ws = wb2.add_sheet(sheet_list_merge[index].decode('utf-8'))
        row = 3

        location_file = open('location.txt', 'r')
        location_line = location_file.readline()
        while location_line:
            name = location_line.split('+')[0]
            #找到相应sheet的位置信息，写入
            if name in sheet_list_merge[index]:
                date = location_line.split('+')[1]
                time_start = location_line.split('+')[2]
                location_start = location_line.split('+')[3]
                time_end = location_line.split('+')[4]
                location_end = location_line.split('+')[5]
                work_time = calc_work_time(time_end, time_start)

                if row == 3:
                    ws.write_merge(0, 0, 0, 5, 'GPS行程简表'.decode('utf-8'))
                    ws.write_merge(1, 1, 0, 0, '姓名:'.decode('utf-8'))
                    ws.write_merge(1, 1, 1, 1, name.decode('utf-8'))
                    ws.write_merge(2, 2, 0, 0, '日期'.decode('utf-8'))
                    ws.write_merge(2, 2, 1, 1, '出发地点'.decode('utf-8'))
                    ws.write_merge(2, 2, 2, 2, '拜访客户/终端'.decode('utf-8'))
                    ws.write_merge(2, 2, 3, 3, '收工地点'.decode('utf-8'))
                    ws.write_merge(2, 2, 4, 4, '出差天数'.decode('utf-8'))
                    ws.write_merge(2, 2, 5, 5, '备注'.decode('utf-8'))
                ws.write_merge(row, row, 0, 0, date)
                ws.write_merge(row, row, 1, 1, location_start.decode('utf-8'))
                #ws.write_merge(row, row, 2, 2, '拜访客户/终端')
                ws.write_merge(row, row, 3, 3, location_end.decode('utf-8'))
                ws.write_merge(row, row, 4, 4, work_time)
                #ws.write_merge(row, row, 5, 5, '备注')
                row += 1
            location_line = location_file.readline()

        ws.write_merge(row+1, row+1, 0, 0, '合计'.decode('utf-8'))
        location_file.close()

    print 'The "report_jicheng.xls" has been created/updated'
    wb2.save(file_name2)

if __name__ == '__main__':
    main()