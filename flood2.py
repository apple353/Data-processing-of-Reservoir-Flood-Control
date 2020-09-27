import pandas as pd 
import numpy as np
import string 
import datetime
df_1 = pd.read_excel(r'C:\Users\ChenJL\Desktop\2018第1次调洪1.xls',sheet_name='tiaohong')
def is_datetime(x):
    if  isinstance(x,str): 
        l = x.split("日")
        l1 = l[1]
        l1 = l1.strip()
        k = pd.to_datetime(l1)
        return k
    else:
        return x
y = eval(input("年"))
m =  eval(input("月"))
d =  eval(input("日"))
size = df_1.iloc[:,0].size
x = list()
for i in range(0,size):
    aa = (df_1.iloc[i,0])
    k = is_datetime(aa)
    if i > 0 :   
        h1 = k.hour
        m1 = k.minute
        bb = df_1.iloc[i-1,0]
        k2 = is_datetime(bb)      
        h2 = k2.hour
        m2 = k2.minute
        if h1 < h2 or isinstance(aa,str):
            d += 1
            if  d>30 and m in (4,6,9,11):
                   m += 1
                   d =0
                   d += 1
                   mydate2 = datetime.datetime( y,m,d,h1, m1, 00)
                   x.append(mydate2) 
            elif d>28 and m ==2 :
                m += 1
                d = 0
                mydate2 = datetime.datetime( y,m,d,h1, m1, 00)
                x.append(mydate2)
            elif d>31:
                m += 1
                d =0
                d += 1
                mydate2 = datetime.datetime( y,m,d,h1, m1, 00)
                x.append(mydate2) 
                     
            else:
                mydate2 = datetime.datetime( y,m,d,h1, m1, 00)
                x.append(mydate2)
                
        else:
            h3 = k.hour
            m3 = k.minute
            mydate3 = datetime.datetime( y,m,d,h3, m3, 00)
            x.append(mydate3)
    else:
        h4 = k.hour
        m4 = k.minute
        mydate4 = datetime.datetime( y,m,d,h4, m4, 00)
        x.append(mydate4)
df_1["时间"]=x
col_name = df_1.columns.tolist()
df = pd.DataFrame(columns= col_name)
for j in range(0,size-1):
    diff = df_1.iloc[j+1,0] - df_1.iloc[j,0]
    level_diff =  df_1.iloc[j+1,2] - df_1.iloc[j,2]
    sec = diff.total_seconds()
    num = int(sec/3600)
    df = df.append(df_1.iloc[[j]],ignore_index=True)
    if sec > 3600:
        time_temp = df_1.iloc[j,0].minute       
        df_temp = df_1.iloc[[j]]
        if time_temp == 00:
            for i in range(0,num): 
                df_temp['时间'] += datetime.timedelta(hours=1) 
                df_temp['时段差值'] = ''
                df_temp['水位'] = round(level_diff/num*(i+1)+df_1.iloc[j,2],2)
                ser = round(level_diff/num*(i+1)+df_1.iloc[j,2],2)
                if ser> 38 :
                    df_temp['库容'] = round((ser-38)*87.83+1444.19,2)
                elif 37< ser < 38:
                    df_temp['库容'] = round((ser-37)*84.03+1360.16,2)
                elif 36<ser<37:
                    df_temp['库容'] = round((ser-36)*80.36+1279.8,2)
                elif 35<ser<36:
                    df_temp['库容'] = round((ser-35)*76.8+1203,2)
                elif 34<ser<35:
                    df_temp['库容'] = round((ser-34)*73.41+1129.59,2)
                else:
                    pass    
                df = df.append(df_temp,ignore_index=True)
        else:
            for i in range(0,num):
                df_temp['时间'] += datetime.timedelta(minutes=30) 
                df_temp['时段差值'] = ''
                df_temp['水位'] = round(level_diff/num*(i+1)+df_1.iloc[j,2],2)
                ser = round(level_diff/num*i+df_1.iloc[j,2],0)
                if ser > 38 :
                    df_temp['库容'] = round((ser-38)*87.83+1444.19,2)
                elif 37< ser < 38:
                    df_temp['库容'] = round((ser-37)*84.03+1360.16,2)
                elif 36<ser<37:
                    df_temp['库容'] = round((ser-36)*80.36+1279.8,2)
                elif 35<ser<36:
                    df_temp['库容'] = round((ser-35)*76.8+1203,2)
                elif 34<ser<35:
                    df_temp['库容'] = round((ser-34)*73.41+1129.59,2)
                else:
                    pass    
                df = df.append(df_temp,ignore_index=True)
df = df.append(df_1.iloc[[size-1]],ignore_index=True)          
size_new = df.iloc[:,0].size
#计算时段差值
s = [0]
for k in range(0,size_new-1):
    time_temp = df.iloc[k+1,0] - df.iloc[k,0]
    secs = time_temp.total_seconds()
    s.append(secs)
df['时段差值']=s
#计算库容差值
re = [0]
for q in range(1,size_new):
    re_temp = round(df.iloc[q,3] - df.iloc[q-1,3],2)
    re.append(re_temp)
df['库容差值'] = re
#计算入库流量
res_q = [0]
for e in range(1,size_new):
    qq_temp = round(df.iloc[e,4]*10000/df.iloc[e,1]+ df.iloc[e-1,8],2)
    res_q.append(qq_temp)
df['入库流量'] = res_q

col_name1 = df_1.columns.tolist()
df_2 = pd.DataFrame(columns= col_name1)
df_2 = df_2.append(df.iloc[[0]])
for row_id in range(1,size_new-1):

    last_t = df.iloc[row_id-1,0]
    next_t = df.iloc[row_id,0]
    df2_temp = df.iloc[[row_id]]
    if last_t == next_t:
        pass
    elif last_t.minute != 00:
        pass
    else:
        df_2 = df_2.append(df2_temp)
df_2 = df_2.append(df.iloc[[size_new-1]],ignore_index=True) 
#保存生成的文件
writer = pd.ExcelWriter(r'C:\Users\ChenJL\Desktop\output.xlsx')
df_2.to_excel(excel_writer=writer,sheet_name='2019年')
df.to_excel(excel_writer=writer,sheet_name='2019年1')
writer.save()
writer.close()


