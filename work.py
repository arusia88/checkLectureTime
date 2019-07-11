#!/usr/bin/python3.4
# -*- coding: utf-8 -*-
#
# install library : pandas, xlrd, openpyxl
# coded with python 3.6
#
# 수강시작과 수강끝 시간은 14자리로 항상 이루어진다고 가정한다.

import pandas as pd
import sys
import os
import copy
import math
import pprint
import traceback
from optparse import OptionParser
from datetime import datetime, timedelta
from collections import OrderedDict

# 추출할 파일명
targetFile = "sample.xlsx"
dataFile = "data.xlsx"
getFileName = lambda name : "./" + name

# 전체 수강기간날짜 선언
startDateForTerm = "201903191700";
endDateForTerm = "201906111430";

# set working directory

os.chdir('./')
parser = OptionParser()
parser.add_option('-s', '--start',
    action='store',
    dest='startDateForTerm',
    help='set startdateForTerm variable. e.g)201903110000')
parser.add_option('-e', '--end',
    action='store',
    dest='endDateForTerm',
    help='set enddateForTerm variable. e.g)201904110000')
parser.add_option('-t', '--target',
    action='store',
    dest='targetFile',
    help='set targetFileName. target file have to locate same directory. default targetFileName is "sample.xlsx" e.g) sample.xlsx')
parser.add_option('-d', '--data',
    action='store',
    dest='dataFile',
    help='set dataFileName. target file have to locate same directory. default dataFileName is "data.xlsx"')

(options, args) = parser.parse_args(sys.argv)

def padding(str):
    t = str + "00000000000000"
    return t[0:14];

targetFile = options.targetFile or targetFile
dataFile = options.dataFile or dataFile
startDateForTerm = padding(options.startDateForTerm or startDateForTerm)
endDateForTerm = padding(options.endDateForTerm or endDateForTerm)

def makeKey(row):
    return str(row['id']) + '_' + row['title'];

def makeData(row):
    t = {'timeData': [], 'ipData': []};
    t['id'] = row['id'];
    t['cid'] = row['cid'];
    t['content'] = row['title'];
    t['week'] = row['week']
    t['accessDevice'] = row['접속기기']
    t['timeData'].append([row['수강시작'], row['수강끝']]);
    t['ipData'].append(row['등록 ip'])
    return t;

# row에서 필요한 부분만 파싱한다.
def parseRow(row, res):
    key = makeKey(row);
    if (key in res):
        res[key]['timeData'].append([row['수강시작'], row['수강끝']]);
        if (row['등록 ip'] in res[key]['ipData']):
            return
        res[key]['ipData'].append(row['등록 ip']);
    else:
        res[key] = makeData(row);

def parseDF(df):
    res = {}
    # loop row
    for i in range(0, len(df)):
        row = df.iloc[i, :];
        parseRow(row, res);
    return res;

# 30분마다 해당 값을 1로 체크 하는 함수
def checkTimePerHalfHour(start, end, target, value):
    gs = datetime(int(startDateForTerm[0:4]), int(startDateForTerm[4:6]), int(startDateForTerm[6:8]), int(startDateForTerm[8:10]), int(startDateForTerm[10:12]))
    s = datetime(int(start[0:4]), int(start[4:6]), int(start[6:8]), int(start[8:10]), int(start[10:12]))
    start_date = s if s >= gs else gs
    ge = datetime(int(endDateForTerm[0:4]), int(endDateForTerm[4:6]), int(endDateForTerm[6:8]), int(endDateForTerm[8:10]), int(endDateForTerm[10:12]))
    e = datetime(int(end[0:4]), int(end[4:6]), int(end[6:8]), int(end[8:10]), int(end[10:12]))
    end_date = ge if ge <= e else e
    d = start_date
    # 30분 마다 키에 value 를 체크한다
    if (value == 0):
        delta = timedelta(seconds=1800)
        while d <= end_date:
            # key값을 조정한다 V|월|일|시|분|전/후
            postfix = '01' if d.strftime("%M") == '00' else '02'
            key = 'V'+d.strftime("%m%d%H")+postfix;
            target[key] = value;
            d += delta;
    else:
        postfix = '01' if d.strftime("%M") == '00' else '02'
        key = 'V'+d.strftime("%m%d%H")+postfix;
        target[key] = value;

# 60초 미만인 경우 검출
def isLessOneMinute(start, end):
    start_date = datetime(int(start[0:4]), int(start[4:6]), int(start[6:8]), int(start[8:10]), int(start[10:12]), int(start[12:14]))
    end_date = datetime(int(end[0:4]), int(end[4:6]), int(end[6:8]), int(end[8:10]), int(end[10:12]), int(end[12:14]))
    diff = end_date - start_date
    return diff.seconds <= 60

def adjustTime(time):
    d = int(time) % 10000
    # 30분보다 큰 경우 분을 3000 으로 작은 경우 0000으로 초기화
    if d >= 3000:
        return time[:-4] + "3000";
    else:
        return time[:-4] + "0000";

def checkCriteria(timeData, cid, ref):
    cSeconds = ref[ref.cid == cid].duration.values[0] * 60
    sumSeconds = 0
    for time in timeData:
        # 둘중 하나라도 빈 값인 경우 패스. 언제까지 체크해야할지 알수없음.
        if math.isnan(time[0]) or math.isnan(time[1]):
            continue
        # 60초 미만인 경우 체크안함
        start = padding(str(int(time[0])));
        end = padding(str(int(time[1])));
        gs = datetime(int(startDateForTerm[0:4]), int(startDateForTerm[4:6]), int(startDateForTerm[6:8]),
                      int(startDateForTerm[8:10]), int(startDateForTerm[10:12]))
        s = datetime(int(start[0:4]), int(start[4:6]), int(start[6:8]), int(start[8:10]), int(start[10:12]))
        start_date = s if s >= gs else gs
        ge = datetime(int(endDateForTerm[0:4]), int(endDateForTerm[4:6]), int(endDateForTerm[6:8]),
                      int(endDateForTerm[8:10]), int(endDateForTerm[10:12]))
        e = datetime(int(end[0:4]), int(end[4:6]), int(end[6:8]), int(end[8:10]), int(end[10:12]))
        end_date = ge if ge <= e else e
        diff = end_date-start_date;
        sumSeconds =sumSeconds + diff.total_seconds()
    # print('criteria : ', (cSeconds * 0.9), 'total time :', sumSeconds);
    return (cSeconds * 0.9) <= sumSeconds


def convertData(dict, ref):
    # 최종 dataframe의 데이터가 담길 r 선언
    r = [];

    # template로 사용할 date객체 생성한다.
    templateDate = {}
    checkTimePerHalfHour(startDateForTerm, endDateForTerm, templateDate, 0);

    # 사전에 파싱했던 결과를 실제 dataframe형태로 변환한다.
    for item in dict.values():
        temp = copy.deepcopy(templateDate);
        # duration시간을 가지고와서 전체수강시간 * 0.9 보다 높은 경우인지 확인
        if checkCriteria(item['timeData'], item['cid'], ref) == False:
            continue;
        # print(item['id']," is success : ", item['content']);
        for time in item['timeData']:
            # 둘중 하나라도 빈 값인 경우 패스. 언제까지 체크해야할지 알수없음.
            if math.isnan(time[0]) or math.isnan(time[1]):
                continue;
            # 60초 미만인 경우 체크안함
            ts = padding(str(int(time[0])));
            te = padding(str(int(time[1])));
            if isLessOneMinute(ts, te):
                continue;
            s = adjustTime(ts);
            e = adjustTime(te);
            checkTimePerHalfHour(s, e, temp, 1);
        temp['id'] = item['id'];
        temp['cid'] = item['cid'];
        temp['ipData'] = item['ipData'];
        temp['ipCount'] = len(item['ipData']);
        temp['content'] = item['content'];
        temp['week'] = item['week']
        temp['accessDevice'] = item['accessDevice']

        # id content가 앞으로 나오도록 정렬한다
        sorted_dict = OrderedDict(sorted(temp.items(), key=sortFn));
        r.append(sorted_dict);

    return sorted(r, key=lambda item: item['id'])
def sortFn(item):
    if item[0] == 'id':
        return "0"
    if item[0] == 'content':
        return "1"
    if item[0] == 'week':
        return "2"
    if item[0] == 'accessDevice':
        return "3"
    if item[0] == 'cid':
        return "4"
    if item[0] == 'ipData':
        return "5"
    if item[0] == 'ipCount':
        return "6"
    return item[0]

def main():

    try:
        # Parsing data, return dictionary
        df = pd.read_excel(getFileName(targetFile), 0, 0);
        ref = pd.read_excel(getFileName(dataFile), 0, 0);
        print("parsingData...")
        dict = parseDF(df);

        # Result
        print("checkingExcel...")
        res = convertData(dict, ref);
        # pprint.pprint(res);

        # Export excel
        print("writingExcel...")
        output = pd.DataFrame(res);
        output.to_excel('output.xlsx')
        print("Completed!")
    except Exception as e:
        print("An error occured")
        print (traceback.format_exc())


main();