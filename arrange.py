#coding:utf-8

import xlrd
import xlwt
from pymongo import MongoClient
import pymongo
import random
import datetime

def get_subject_type(name):
    if name=='语文':
        return 'Chi'
    if name=='理科数学':
        return 'SMa'
    if name=='文科数学':
        return 'AMa'
    if name=='英语':
        return 'Eng'
    if name=='物理':
        return 'Phy'
    if name=='化学':
        return 'Che'
    if name=='生物':
        return 'Bio'
    if name=='政治':
        return 'Phi'
    if name=='地理':
        return 'Geo'
    if name=='历史':
        return 'His'


dbconn = MongoClient('127.0.0.1', 27017)
db = dbconn['lecture_arrangement']
lecture_col = db['lecture_data']
grade_col = db['grade_data']
lecture_data = xlrd.open_workbook('课程安排数据.xlsx')
grade_data = xlrd.open_workbook('成绩数据.xls')
lecture_sheet = lecture_data.sheets()[0]

lecture_col.delete_many({})
for i in range(1,lecture_sheet.nrows):
    data=lecture_sheet.row_values(i)
    #print(data)
    new_post={
        'name':data[0],
        'subject':data[1],
        'level':data[2],
        'time':data[3],
        'amount':data[4],
        'count':0,
        'list':[]
    }
    lecture_col.insert_one(new_post)
for item in lecture_col.find():
    #print(item)
    pass

grade_col.delete_many({})
for i in grade_col.find():
    #print(i)
    pass


for grade_sheet in grade_data.sheets():
    subject = grade_sheet.cell(0,1).value
    #print(subject)
    count=0
    for i in range(2, grade_sheet.nrows):
        data = grade_sheet.row_values(i)
        #print(grade_col.find({'class': data[1],'name': data[2]}) is True)
        if grade_col.count({'class': data[1],'name': data[2]}):
            print('update',{'class': data[1],'name': data[2]})
            grade_col.update_one({'class': data[1],'name': data[2]},{'$set': {get_subject_type(subject): {
                'rank': data[0],
                'score': data[3]
            }}})
        else:
            print('insert',{'class': data[1],'name': data[2]})
            grade_col.insert_one({'class': data[1],
            'name': data[2],
            get_subject_type(subject): {
                'rank': data[0],
                'score': data[3]
            },
                                  'list':['Empty','Empty','Empty','Empty','Empty','Empty'],
                                  'fail_list':[]})

'''
#生物排课
students_of_bio = grade_col.find({'Bio':{'$exists':True}},projection={'class':True,'name':True,'Bio':True,'list':True}).sort([('Bio.rank', pymongo.ASCENDING)])
lecture_of_bio = lecture_col.find({'subject': 'Bio'}).sort([('level',pymongo.ASCENDING)])
print('课表')
for item in  lecture_col.find({'subject': 'Bio'}).sort([('level',pymongo.ASCENDING)]):
    print(item)
for student in students_of_bio:
    lecture_of_bio = lecture_col.find({'subject': 'Bio'}).sort([('level', pymongo.ASCENDING)])
    for lecture in lecture_of_bio:
        if lecture['count'] < lecture['amount'] and (student['list'][int(lecture['time'])-1]=='Empty'):
            student['list'][int(lecture['time'])-1]=lecture['name']
            lecture['count']+=1
            lecture['list'].append({'class':student['class'],'name':student['name']})
            grade_col.update_one({'class':student['class'],'name':student['name']},{'$set':{'list':student['list']}})
            lecture_col.update_one({'name':lecture['name']},{'$set':{'count':lecture['count'],'list': lecture['list']}})
            print('已为学生 ',student['name'],'安排',lecture['name'])
            break
for i in grade_col.find():
    print(i)
for item in lecture_col.find():
    #print(item)
    pass

#物理排课
students_of_Phy = grade_col.find({'Phy':{'$exists':True}},projection={'class':True,'name':True,'Phy':True,'list':True}).sort([('Phy.rank', pymongo.ASCENDING)])
lecture_of_Phy = lecture_col.find({'subject': 'Phy'}).sort([('level',pymongo.ASCENDING)])
print('课表')
for item in  lecture_col.find({'subject': 'Phy'}).sort([('level',pymongo.ASCENDING)]):
    print(item)
for student in students_of_Phy:
    lecture_of_Phy = lecture_col.find({'subject': 'Phy'}).sort([('level', pymongo.ASCENDING)])
    for lecture in lecture_of_Phy:
        if lecture['count'] < lecture['amount'] and (student['list'][int(lecture['time'])-1]=='Empty'):
            student['list'][int(lecture['time'])-1]=lecture['name']
            lecture['count']+=1
            lecture['list'].append({'class':student['class'],'name':student['name']})
            grade_col.update_one({'class':student['class'],'name':student['name']},{'$set':{'list':student['list']}})
            lecture_col.update_one({'name':lecture['name']},{'$set':{'count':lecture['count'],'list': lecture['list']}})
            print('已为学生 ',student['name'],'安排',lecture['name'])
            break
'''
def arrange(subject,key):
    students_of_Phy = grade_col.find({subject:{'$exists':True}},projection={'class':True,'name':True,subject:True,'list':True}).sort([(subject+'.rank', pymongo.ASCENDING)])
    lecture_of_Phy = lecture_col.find({'subject': subject}).sort([('level',pymongo.ASCENDING)])
    amount=[0,0,0,0]
    count=[0,0,0,0]
    for level in range(1,5):
        for item in lecture_col.find({'subject':subject,'level':level}):
            amount[level-1] += item['amount']
    turn = key
    for student in students_of_Phy:
        level=1 #查询应该安排到哪个级别
        while count[level-1]>=amount[level-1]:
            level+=1
        lecture_of_Phy = lecture_col.find({'subject': subject,'level':level}).sort([('level', pymongo.ASCENDING)])
        number_of_suitable=lecture_col.count({'subject': subject,'level':level})
        try_alert=0
        turn = random.randint(0, number_of_suitable - 1)

        while count[level-1]<=amount[level-1]:
            try_alert+=1
            if try_alert>=10:
                print('由于',student['name'],lecture['name'],lecture['count'],'排课失败')
                fail_list=grade_col.find_one({'class': student['class'], 'name': student['name']})['fail_list']
                fail_list.append(lecture['name'])
                grade_col.update_one({'class': student['class'], 'name': student['name']},
                                     {'$set': {'fail_list':fail_list }})
                break
            turn += 1
            if turn >= number_of_suitable:
                turn = 0

            print(amount)
            print(count)
            lecture=lecture_of_Phy[turn]
            #print('尝试排课',student['name'],lecture['name'], turn, level, number_of_suitable)
            if lecture['count'] < lecture['amount'] and (student['list'][int(lecture['time'])-1]=='Empty'):
                student['list'][int(lecture['time'])-1]=lecture['name']
                lecture['count']+=1
                lecture['list'].append({'class':student['class'],'name':student['name'],'score':student[subject]['score']})
                grade_col.update_one({'class':student['class'],'name':student['name']},{'$set':{'list':student['list']}})
                lecture_col.update_one({'name':lecture['name']},{'$set':{'count':lecture['count'],'list': lecture['list']}})
                print('已为学生 ',student['name'],'安排',lecture['name'],turn)
                count[level-1]+=1
                break
    return True
subjects = ['Bio','Phy','SMa','Che','Geo','AMa','Chi','Eng','His','Phi']
success_flag=[False,False,False,False,False,False,False,False,False,False]
def clear():
    grade_col.update_many({},{'$set':{'list':['Empty','Empty','Empty','Empty','Empty','Empty']}})
    lecture_col.update_many({},{'$set':{'count':0,'list':[]}})

def attemp(subjects,i=0,key=0):
    for test in range(0,10):
        if success_flag[test] is True:
            continue
        else:
            break
        print('成功')
        return
    if arrange(subjects[i],key):
        success_flag[i]=True
        if i+1>=10:
            print('成功')
            return
        attemp(subjects,i+1,key)

    else:
        success_flag[i]=False
        clear()
        attemp(subjects,0,key)

attemp(subjects)


for i in grade_col.find():
    print(i)
for item in lecture_col.find():
    print(item)
    pass

student_result=xlwt.Workbook()
student_sheet=student_result.add_sheet(u'学生课表',cell_overwrite_ok=True)
row0=[u'班级',u'姓名',u'1',u'2',u'3',u'4',u'5',u'6']
for i in range(0,len(row0)):
    student_sheet.write(0,i,row0[i])
counter=1
for item in grade_col.find():
    student_sheet.write(counter,0,item['class'])
    student_sheet.write(counter, 1, item['name'])
    for j in range(0,6):
        student_sheet.write(counter, 2+j, item['list'][j])
    k=0
    while k < len(item['fail_list']):
        student_sheet.write(counter, 9+k, item['fail_list'][k])
        k+=1
    counter+=1
student_result.save('Student_result.xls')

lecture_result=xlwt.Workbook()
for item in lecture_col.find():
    sheet=lecture_result.add_sheet(item['name'])
    counter=1
    sheet.write(0, 0,u'班级')
    sheet.write(0, 1, u'姓名')
    sheet.write(0, 2, u'成绩')
    for k in item['list']:
        sheet.write(counter,0,k['class'])
        sheet.write(counter, 1, k['name'])
        sheet.write(counter, 2, k['score'])
        counter+=1
lecture_result.save('Lecture_result.xls')
