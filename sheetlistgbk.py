__author__ = 'willard'
#-*- coding: gbk -*-

#���ܣ���ģ���ļ��л�ȡsheet�б�������б�
#sheet�б���������Ŀ���ļ��ĸ���sheet
#�����û��ļ�����
#���ģ���ļ��и��£���Ҫ�������б����򣬸����б��ļ�

from xlrd import open_workbook

def main():
    print '��ȷ��ģ���ļ���Ϊmodel.xls�����ͱ��ļ���ͬһ���ļ�����'
    print '**********����������**********'
    print '��ʼ����...'

    BauModel = open_workbook('model.xls')
    dest = open('sheetlist.txt', 'w')
    dest1 = open('namelist.txt', 'w')
    Number = 0
    print '���ڻ�ȡ��Ϣ�����Ժ�...'

    for s in BauModel.sheets():
        if ((Number > 0) & (Number%2 == 0)):
            if (Number != 2):
                dest1.write('\n')
            dest1.write(s.name.encode('gbk')[:-6])


        if (Number%2 ==1):
            if (Number >= 2):
                dest.write('\t')
            dest.write(s.name.encode('gbk'))


        Number += 1

    dest.close()
    dest1.close()
    print 'sheet�ѱ�����sheetlist.txt��'
    print '�����Ѿ�������namelist.txt�ļ���'
    print '**********�������˳�**********'

if __name__ == '__main__':
    main()