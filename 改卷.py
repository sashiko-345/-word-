import pandas as pd
import random
import numpy as np
def check(n):
    stu_ans  = pd.read_excel("提交答案-%d.xlsx"%n, sheet_name=0,usecols=[0],header=None)
    stu_ans = stu_ans.values.tolist()
    ans = open("answers-%d.txt"%n , 'r+')
    cont = ans.read()
    cont = eval(cont)
    answ = []
    for i in range(len(cont)):
        for j in range(len(cont[i])):
            answ.append(cont[i][j])
    score = []
    for i in range(len(answ)):
        if answ[i]==stu_ans[i][0]:
            score.append(1)
        else:
            score.append(0)
    return score
def score(score,m1,m2,m3,n1,n2,n3):
    mark=0
    for i in (0,n1):
        mark += score[i] * m1
    for i in (n1,n2):
        mark += score[i] * m2
    for i in (n2,n3):
        mark += score[i] * m3
    return mark

re = open('record.txt','r+')
r=re.read()
r=eval(r)
print("输入每种题型的分值：")
m1 = int(input('单选题：'))
m2 = int(input('多选题：'))
m3 = int(input('判断题：'))
s=[]
for i in range(r[0]):
    s.append(score(check(i),m1,m2,m3,r[1]+r[2]+r[3],r[4]+r[5]+r[6],r[7]+r[8]+r[9]))

for i in range(r[0]):
    print("%d的分数是%d"%(i,s[i]))

print("全班平均分是%d",np.mean(s))
p = int(input("求前百分之n的分数线？请输入n"))
s.sort()
print(s[int(p*len(s)/100)])





