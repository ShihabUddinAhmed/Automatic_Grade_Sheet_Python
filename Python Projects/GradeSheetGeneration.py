import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
q1df = pd. read_excel (r'Quizes/Quiz 1.xlsx', sheet_name='Sheet1')
sid1=np.array(len(q1df.index), dtype='str')
sid1=pd.DataFrame(q1df, columns=['Email']).to_numpy()
for i in range(0,len(q1df.index)):
    sid1[i][0]=sid1[i][0].split('@')
    sid1[i]=sid1[i][0][0]
score1=np.array(len(q1df.index), dtype=np.int32)
score1=pd.DataFrame(q1df, columns=['Total points']).to_numpy()


q2df = pd. read_excel (r'Quizes/Quiz 2.xlsx', sheet_name='Sheet1')
sid2=np.array(len(q2df.index), dtype='str')
sid2=pd.DataFrame(q2df, columns=['Email']).to_numpy()
for i in range(0,len(q2df.index)):
    sid2[i][0]=sid2[i][0].split('@')
    sid2[i]=sid2[i][0][0]
score2=np.array(len(q2df.index), dtype=np.int32)
score2=pd.DataFrame(q2df, columns=['Total points']).to_numpy()


q3df = pd. read_excel (r'Quizes/Quiz 3.xlsx', sheet_name='Sheet1')
sid3=np.array(len(q3df.index), dtype='str')
sid3=pd.DataFrame(q3df, columns=['Email']).to_numpy()
for i in range(0,len(q3df.index)):
    sid3[i][0]=sid3[i][0].split('@')
    sid3[i]=sid3[i][0][0]
score3=np.array(len(q3df.index), dtype=np.int32)
score3=pd.DataFrame(q3df, columns=['Total points']).to_numpy()


ldf = pd. read_excel (r'Lab Exam.xlsx', sheet_name='Sheet1')
sid4=np.array(len(ldf.index), dtype='str')
sid4=pd.DataFrame(ldf, columns=['Email']).to_numpy()
for i in range(0,len(ldf.index)):
    sid4[i][0]=sid4[i][0].split('@')
    sid4[i]=sid4[i][0][0]
score4=np.array(len(ldf.index), dtype=np.int32)
score4=pd.DataFrame(ldf, columns=['Total points']).to_numpy()


adf = pd. read_csv ('Assignment.csv')
sid5=np.array(len(adf.index), dtype='str')
sid5=pd.DataFrame(adf, columns=['Student ID']).to_numpy()
score5=np.array(len(adf.index), dtype=np.int32)
score5=pd.DataFrame(adf, columns=['Ass.']).to_numpy()
name=np.array(len(adf.index), dtype='str')
name=pd.DataFrame(adf, columns=['Name']).to_numpy()
space=" "
sname=" "
for i in range(0,len(adf.index)):
    name[i][0]=name[i][0].split(', ')
    if len(name[i][0])==2:
        sname=np.char.add(name[i][0][1],space)
        name[i]=np.char.add(sname,name[i][0][0])
    else:
        name[i]=name[i][0][0]


file_data = open('Attendance_files/Week 1 Lab .csv')
name0=""
rows=[100]
i=0
for row in file_data:
    row=row.split('\x00')
    rows=row
    for x in rows:
        name0+=x
    i+=1
name0=name0.split('\t')
i=0
for x in name0:
    name0[i]=x.split('\n')
    if len(name0[i])==2:
        name0[i]=name0[i][1]
    else:
        name0[i]="Empty"
    i+=1

file_data = open('Attendance_files/Week 1 Theory.csv')
name1=""
rows=[100]
i=0
for row in file_data:
    row=row.split('\x00')
    rows=row
    for x in rows:
        name1+=x
    i+=1
name1=name1.split('\t')
i=0
for x in name1:
    name1[i]=x.split('\n')
    if len(name1[i])==2:
        name1[i]=name1[i][1]
    else:
        name1[i]="Empty"
    i+=1


file_data = open('Attendance_files/Week 2 Theory.csv')
name2=""
rows=[100]
i=0
for row in file_data:
    row=row.split('\x00')
    rows=row
    for x in rows:
        name2+=x
    i+=1
name2=name2.split('\t')
i=0
for x in name2:
    name2[i]=x.split('\n')
    if len(name2[i])==2:
        name2[i]=name2[i][1]
    else:
        name2[i]="Empty"
    i+=1


file_data = open('Attendance_files/Week 4 Lab (Makeup).csv')
name3=""
rows=[100]
i=0
for row in file_data:
    row=row.split('\x00')
    rows=row
    for x in rows:
        name3+=x
    i+=1
name3=name3.split('\t')
i=0
for x in name3:
    name3[i]=x.split('\n')
    if len(name3[i])==2:
        name3[i]=name3[i][1]
    else:
        name3[i]="Empty"
    i+=1


file_data = open('Attendance_files/Week 5 Lab.csv')
name4=""
rows=[100]
i=0
for row in file_data:
    row=row.split('\x00')
    rows=row
    for x in rows:
        name4+=x
    i+=1
name4=name4.split('\t')
i=0
for x in name4:
    name4[i]=x.split('\n')
    if len(name4[i])==2:
        name4[i]=name4[i][1]
    else:
        name4[i]="Empty"
    i+=1
rows, cols = (len(sid5), 16)
result=[]
for i in range(rows):
    col = []
    for j in range(cols):
        col.append(0)
    result.append(col)

for r in range(0,len(sid5)):
    result[r][0]=sid5[r][0]
    result[r][1]=name[r][0]
    result[r][13]=score5[r][0]
    for q in range(0,len(sid1)):
        if(sid1[q][0]==result[r][0]):
            result[r][8]=score1[q][0]
        else:
            continue
    for q in range(0,len(sid2)):
        if(sid2[q][0]==result[r][0]):
            result[r][9]=score2[q][0]
        else:
            continue
    for q in range(0,len(sid3)):
        if(sid3[q][0]==result[r][0]):
            result[r][10]=score3[q][0]
        else:
            continue
    for q in range(0,len(sid4)):
        if(sid4[q][0]==result[r][0]):
            result[r][12]=score4[q][0]
        else:
            continue
    for nam in range(0,len(name0)):
        if(name0[nam]==result[r][1]):
            result[r][2]=1
        else:
            continue
    for nam in range(0,len(name1)):
        if(name1[nam]==result[r][1]):
            result[r][3]=1
        else:
            continue
    for nam in range(0,len(name2)):
        if(name2[nam]==result[r][1]):
            result[r][4]=1
        else:
            continue
    for nam in range(0,len(name3)):
        if(name3[nam]==result[r][1]):
            result[r][5]=1
        else:
            continue
    for nam in range(0,len(name4)):
        if(name4[nam]==result[r][1]):
            result[r][6]=1
        else:
            continue
for r in range(0,len(sid5)):
    result[r][7]=(result[r][2]+result[r][3]+result[r][4]+result[r][5]+result[r][6])*2
    list1=[result[r][8],result[r][9],result[r][10]]
    list1.sort()
    result[r][11]=list1[-1]+list1[-2]
    result[r][14]=result[r][13]+result[r][12]+result[r][11]+result[r][7]
    if result[r][14]>=90:
        result[r][15]='A+'
    elif result[r][14]>=85:
        result[r][15]='A'
    elif result[r][14]>=80:
        result[r][15]='B+'
    elif result[r][14]>=75:
        result[r][15]='B'
    elif result[r][14]>=70:
        result[r][15]='C+'
    elif result[r][14]>=65:
        result[r][15]='C'
    elif result[r][14]>=60:
        result[r][15]='D+'
    elif result[r][14]>=50:
        result[r][15]='D'
    else:
        result[r][15]='F'
df = pd.DataFrame(data=result, columns=["Student ID", "Name", "Week 1 LAB", "Week 1 Theory", "Week 2 Theory", "Week 4 LAB(Makeup)", "Week 5 LAB", "Attendance Marks", "Quiz1", "Quiz2", "Quiz3", "Quiz Total", "Lab Exam", "Assignment", "Total Marks", "Grade"])

df.to_excel (r'FinalGradeSheet.xlsx', index = False, header = True)