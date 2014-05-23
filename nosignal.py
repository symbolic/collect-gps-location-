__author__ = 'willard'
#-*- coding: utf-8 -*-
import os

gap_ignore = 30
#设定的最长离线时间，单位：分钟
#设定的最短在线时间，单位：小时
little_online_time = 4
mark = '\t'


def time_format(number):
    hh = str(number/3600)
    mm = str((number-int(hh)*3600)/60)
    ss = str(number % 60)

    if len(hh) == 1:
        hh = '0' + hh
    if len(mm) == 1:
        mm = '0' + mm
    if len(ss) == 1:
        ss = '0' + ss

    return hh + '小时' + mm + '分' + ss + '秒'


def time_second(time):
    hour = int(time.split(':')[0])
    minute = int(time.split(':')[1])
    second = int(time.split(':')[2])

    return hour*3600+minute*60+second


def time_gap(start, end):
    time_start = start.split(' ')[1]
    time_end = end.split(' ')[1]

    return time_second(time_end)-time_second(time_start)


def handle_online(aline, dest, dest2):
    str1 = aline.split('+')
    if len(str1) > 6:
        print aline
        print 'This line is error [(location.txt)], please check it.\n'
    name = str1[0]
    date = str1[1]
    start_time = str1[2]
    start_location = str1[3]
    end_time = str1[4]

    if start_location != ' ':
        online_second = time_second(end_time) - time_second(start_time)
        dest.write(name + mark + date + mark + start_time + '----' + end_time + mark + time_format(online_second) + '\n')
        if online_second < little_online_time*3600:
            dest2.write(name + mark + date + mark + start_time + '----' + end_time + mark + time_format(online_second) + '\n')


def handle(aline, dest, dest2):
    #全天都是无数据
    if '全天'in aline:
        dest.write(aline)

    #局部无数据
    else:
        str1 = aline.split('\t')
        name = str1[0]
        date_time_start = str1[1]
        location = str1[2]
        date_time_end = str1[3]

        second = time_gap(date_time_start, date_time_end)
        if second > gap_ignore*60:
            dest.write(name + '\t' + date_time_start + '----'
                       + date_time_end + '\t' + location + '\t' + time_format(second) + '\n')
        else:
            dest2.write(name + '\t' + date_time_start + '----'
                        + date_time_end + '\t' + location + '\t' + time_format(second) + '\n')


def run_check(file_name):
    if not os.path.isfile(file_name):
        print 'The file "'+file_name+'" is not exist, please run other file first.'
        return 0


def main():
    print '从no_signal.txt中计算掉线时间'
    print('no_signal2.txt: 掉线时间间隔超过30分钟，最后一列为掉线时间')
    print('no_signal3.txt: 掉线时间间隔不足30分钟，最后一列为掉线时间')
    if run_check('no_signal.txt') == 0:
        return 0

    src = open('no_signal.txt', 'r')
    dest = open('no_signal2.txt', 'w')
    dest2 = open('no_signal3.txt', 'w')

    aline = src.readline()
    while aline:
        handle(aline, dest, dest2)
        aline = src.readline()

    src.close()
    dest.close()
    dest2.close()

    src2 = open('location.txt', 'r')
    dest3 = open('online_time.txt', 'w')
    dest4 = open('online_time_little.txt', 'w')
    aline = src2.readline()
    while aline:
        handle_online(aline, dest3, dest4)
        aline = src2.readline()

    src2.close()
    dest3.close()
    dest4.close()


if __name__ == '__main__':
    main()
