__author__ = 'willard'
#-*- coding: utf-8 -*-
from xlrd import open_workbook
import os
mark = '+'
mark1 = '\t'


#返回当月的天数
def day_of_month(date_time):
    long_month = [1, 3, 5, 7, 8, 10, 12]
    normol_month = [4, 6, 9, 11]

    date = date_time.split(' ')[0]
    year = int(date.split('-')[0])
    month = int(date.split('-')[1])

    if month in long_month:
        return 31
    if month in normol_month:
        return 30
    if (year % 4 == 0) and (year % 100 != 0) or (year % 400 == 0):
        return 29
    else:
        return 28


def calc_date(date_time):
    return date_time.split(' ')[0]


def calc_time(date_time):
    return date_time.split(' ')[1]


#返回时间差（天数）
#<0  未更新
#=0，时间连贯
#>0，时间跳变
def day_gap(date_time_last, date_time_current):
    return day_date_time(date_time_current) - day_date_time(date_time_last)


#生成缺少数据的日期
def create_date(dest, error_file, date_time_last, date_time_current, name, status):
    day_current = day_date_time(date_time_current)
    day_last = day_date_time(date_time_last)

    #status != 1,起始行不是1日，或中间跳变
    #status == 1,最后一行不是最后一天，补充缺少的日期
    if status == 1:
        day_last = day_current
        day_current = day_of_month(date_time_current)+1

    for i in range(day_last, day_current-1):
        if i <= 8:
            new_date = date_time_current[:-11] + '0' + str(i+1)
        else:
            new_date = date_time_current[:-11] + str(i+1)

        dest.write(name + mark + new_date + mark + ' ' + mark + ' ' + mark + ' ' + mark + ' \n')
        error_file.write(name + mark1 + new_date + mark1 + '全天无数据\n')


def day_date_time(date_time):
    date = date_time.split(' ')[0]

    return int(date.split('-')[2])


def pickup(original_excel, dest, error_file):
    #起始地址是否写入，0：没有写入；1：写入。默认起点没有写入
    write_start = 0
    #终点地址是否写入，0：没有写入；1：写入。默认终点已经写入
    write_end = 1
    #用于和第一行的日期比较，时间是否跳变
    date_time_last = '2000-00-00 00:00:00'
    location_last = ''
    always_no_signal = 0
    for row in range(original_excel.nrows):
        #第一行是标题，跳过
        if row == 0:
            continue

        #如果没有位置信息，则跳过
        if len(str(original_excel.cell(row, 3).value.encode('utf-8'))) < 3:
            continue

        #记录当前行数据
        name = str(original_excel.cell(row, 0).value.encode('utf-8'))
        #phoneNum = str(original_excel.cell(row, 1).value.encode('utf-8'))
        date_time_current = str(original_excel.cell(row, 2).value.encode('utf-8'))
        location_current = str(original_excel.cell(row, 3).value.encode('utf-8'))

        #时间不变
        if day_gap(date_time_last, date_time_current) == 0:

            #起点位置已经写入
            if write_start == 1:
                #上一行位置有效，本行位置无效，写入location_current到异常文件，location_middle暂存
                if location_current[:24] == '没有收到卫星信号':
                    always_no_signal = 1
                    if location_last[:24] != '没有收到卫星信号':
                        error_file.write(name + mark1 + date_time_current + mark1 + location_current + mark1)
                        date_time_middle = date_time_last
                        location_middle = location_last

                #上一行位置无效，本行位置有效，上一行写入异常location_last
                if location_last[:24] == '没有收到卫星信号':
                    if location_current[:24] != '没有收到卫星信号':
                        always_no_signal = 0
                        error_file.write(date_time_last + mark1 + location_last + '\n')
            #起点位置没有写入
            if write_start == 0:
                #当前位置有效，写入 location_current
                if location_current[:24] != '没有收到卫星信号':
                    dest.write(name + mark + calc_date(date_time_current) +
                               mark + calc_time(date_time_current) + mark + location_current + mark)
                    write_start = 1
                    write_end = 0
                    if location_last[:24] == '没有收到卫星信号':
                        error_file.write(date_time_last + mark1 + location_last + '\n')

        #时间变化了
        if day_gap(date_time_last, date_time_current) > 0:
            #上一行是无效信号，写入异常文件 date_time_last
            #if location_last[:24] == '没有收到卫星信号':
            #    error_file.write(name + mark1 + date_time_last + mark1 + location_last + '\n')
            #前一条记录终点位置已经写入，只需要处理本行就可以了
            #if write_end == 1:
            #前一条记录终点位置没有写入
            if write_end == 0:
                #上一行的位置是有效位置，则直接写入 location_last
                if location_last[:24] != '没有收到卫星信号':
                    dest.write(calc_time(date_time_last) + mark + location_last + '\n')
                #上一行的位置无效，写入之间保存的有效位置 location_middle
                else:
                    dest.write(calc_time(date_time_middle) + mark + location_middle + '\n')
                    error_file.write(date_time_last + mark1 + location_last + '\n')
                write_end = 1
                write_start = 0
                #本行位置有效，直接写入location_current
            #一整天都是无信号，则写入location_no_signal_start 和 location_last
            elif always_no_signal == 1:
                dest.write(name + mark + calc_date(date_time_last.split(' ')[0]) + mark + ' ' + mark + ' ' + mark + ' ' + mark + ' ' + '\n')
                error_file.write(date_time_last + mark1 + location_last + '\t全天无有效数据\n')
            #日期跳变了，补充上
            if day_gap(date_time_last, date_time_current) > 1:
                create_date(dest, error_file, date_time_last, date_time_current, name, 0)

            if location_current[:24] != '没有收到卫星信号':
                always_no_signal = 0
                dest.write(name + mark + calc_date(date_time_current) + mark + calc_time(date_time_current) + mark + location_current + mark)
                write_start = 1
                write_end = 0
            #本行位置无效，写入异常文件
            else:
                always_no_signal = 1
                date_time_no_signal_start = date_time_current
                location_no_signal_start = location_current
                error_file.write(name + mark1 + date_time_current + mark1 + location_current + mark1)

        date_time_last = date_time_current
        location_last = location_current

    #本行位置有效，直接写入location_current
    if location_current[:24] != '没有收到卫星信号':
        dest.write(calc_time(date_time_current) + mark + location_current + '\n')

    #最后1行，1整天都没有信号
    if (location_current[:24] == '没有收到卫星信号')&(always_no_signal == 1):
        dest.write(name + mark + calc_date(date_time_last.split(' ')[0]) + mark + ' ' + mark + ' ' + mark + ' ' + mark + ' ' + '\n')
        error_file.write(date_time_last + mark1 + location_last + '\t全天无有效数据\n')

    #本行位置无效，直接写入location_middle
    if (location_current[:24] == '没有收到卫星信号')&(always_no_signal != 1):
        dest.write(calc_time(date_time_middle) + mark + location_middle + '\n')

    #补充日期
    create_date(dest, error_file, date_time_last, date_time_current, name, 1)


def run_check(file_name):
    if not os.path.isfile(file_name):
        print 'The file "'+file_name+'" is not exist, please check again.'
        return 0


def main():
    if run_check('beinongda.xls') == 0:
        return 0

    beinongda_file = open_workbook('beinongda.xls')
    dest = open('location.txt', 'w')
    error_file = open('no_signal.txt', 'w')

    for s in beinongda_file.sheets():
        local_name = open('namelist.txt', 'r')
        person_name = local_name.readline()

        while person_name:
            if s.name.encode('utf-8')[:6] == person_name[:6]:
                local_name.close()
                pickup(s, dest, error_file)
                break

            person_name = local_name.readline()

    dest.close()
    error_file.close()

if __name__ == '__main__':
    main()
