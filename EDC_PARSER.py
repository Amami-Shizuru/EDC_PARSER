import os
import pandas as pd

def parseRelation(raw_data):
    print("正在解析关联AE")
    raw_data['prev']=''
    raw_data['next']=''
    raw_data['CTCAE等级']=''
    raw_data['等级变化及时间']=''
    for index, row in raw_data.iterrows():
        for index_y, row_y in raw_data.iterrows():
            match = index_y != index \
                    and row_y['不良事件名称'] == row['不良事件名称'] \
                    and row_y['受试者编号'] == row['受试者编号'] \
                    and row_y['与试验药关系'] == row['与试验药关系'] \
                    and row_y['结束时间'] == row['开始时间'] 

            if match:
                #raw_data.loc[index_y, 'before'] = row['NCI-CTCAE5.0分级']
                #raw_data.loc[index_y, 'after'] = row_y['NCI-CTCAE5.0分级']
                #raw_data.loc[index_y, '变化时间'] = row_y['开始时间']
                raw_data.loc[index_y, 'prev'] = index
                raw_data.loc[index, 'next'] = index_y
    print("关联AE解析完成")
    return raw_data

def mergeRelation(raw_data):
    print("正在合并关联AE")
    for index, row in raw_data.iterrows():
        try:
            if row['prev'] != '' or row['next'] != '':
                cur = index
                ll = [cur]
                while raw_data.loc[cur,'prev'] != '':
                    ll.insert(0,raw_data.loc[cur,'prev'])
                    cur = raw_data.loc[cur, 'prev']
                cur = index
                while raw_data.loc[cur,'next'] != '':
                    ll.append(raw_data.loc[cur,'next'])
                    cur = raw_data.loc[cur, 'next']
                lev_str = ""
                lev_with_date = ""
                for i in range(len(ll)):
                    lev = raw_data.loc[ll[i],'NCI-CTCAE5.0分级'] 
                    lev_str += lev
                    if i < len(ll) - 1:
                        lev_str += "->"
                        lev_next = raw_data.loc[ll[i+1],'NCI-CTCAE5.0分级'] 
                        change_date = raw_data.loc[ll[i+1],'开始时间']
                        lev_with_date += str(lev) + "->" + str(lev_next) + "," + change_date
                    if i < len(ll) - 2:
                        lev_with_date += ";"
                raw_data.loc[ll[0],'CTCAE等级'] = lev_str
                raw_data.loc[ll[0],'等级变化及时间'] = lev_with_date
                print('删除关联条目:',ll[1:])
                raw_data.drop(ll[1:],inplace=True)
        except:
            pass
    for index, row in raw_data.iterrows():
        if row['CTCAE等级'] == '':
          raw_data.loc[index,'CTCAE等级'] = row['NCI-CTCAE5.0分级']
    raw_data.drop(['NCI-CTCAE5.0分级'],axis=1,inplace=True)
    raw_data.drop(['prev'],axis=1,inplace=True)
    raw_data.drop(['next'],axis=1,inplace=True)
    raw_data.drop(['表'],axis=1,inplace=True)
    raw_data.drop(['表名称'],axis=1,inplace=True)
    raw_data.drop(['记录ID'],axis=1,inplace=True)
    raw_data.drop(['受试者信息记录ID'],axis=1,inplace=True)
    raw_data.drop(['项目代码'],axis=1,inplace=True)
    raw_data.drop(['CRF状态'],axis=1,inplace=True)
    raw_data.drop(['研究中心名称'],axis=1,inplace=True)
    raw_data.drop(['研究者姓名'],axis=1,inplace=True)
    raw_data.drop(['研究者代码'],axis=1,inplace=True)
    raw_data.drop(['访视名称'],axis=1,inplace=True)
    raw_data.drop(['访视编号'],axis=1,inplace=True)
    raw_data.drop(['访视内序号'],axis=1,inplace=True)
    raw_data.drop(['eCRF版本号'],axis=1,inplace=True)
    raw_data.drop(['方案版本号'],axis=1,inplace=True)
    raw_data.drop(['自签署知情同意书以来，有无不良事件发生'],axis=1,inplace=True)
    raw_data.drop(['序号'],axis=1,inplace=True)
    raw_data.drop(['患者是否因此不良事件而退出试验'],axis=1,inplace=True)
    raw_data.drop(['是否为SAE，如是请填写SAE表格'],axis=1,inplace=True)
    raw_data.drop(['创 建 人'],axis=1,inplace=True)
    raw_data.drop(['创建单位'],axis=1,inplace=True)
    raw_data.drop(['创建时间'],axis=1,inplace=True)
    raw_data.drop(['修 改 人'],axis=1,inplace=True)
    raw_data.drop(['修改时间'],axis=1,inplace=True)
    print("关联AE合并完成")
    return raw_data


if __name__ == '__main__':
    dir = './data'
    for root, dirs, files in os.walk(dir):
        for f in files:
            if f.endswith(".xlsx"):
                print("正在处理" + f)
                file_path = os.path.join(root, f)
                raw_data = pd.read_excel(file_path, sheet_name='不良事件表(ae)', header=0,keep_default_na=False)
                raw_data = parseRelation(raw_data)
                raw_data = mergeRelation(raw_data)
                output_file = 'output_' + f
                print("保存到" + output_file)
                output_path = './' + output_file
                writer = pd.ExcelWriter(output_path)
                raw_data.to_excel(writer, index=False, sheet_name='不良事件相关检测')
                writer.save()
