import os
import re
import traceback
import matplotlib.pyplot as plt
import pandas as pd
import win32com.client as win32
import shutil


# 将带有xlm的xls文件转换为pandas可以直接读取的xlsx文件
# path：若为'.'，则转换当前目录下；如果为文件夹绝对路径则转换绝对路径下的xls
def toXlsx(path):
    print("开始将xls文件转换为xlsx文件：")
    for name in os.listdir(path):
        # print(name)
        if name.endswith('xls'):
            # 读取当前目录下的xls文件，并将其转换为xlsx文件
            print('读取 '+ name)
            # re.search('[<\>\?\[\]\:\|\*]',name)
            newName = re.sub('[<\>\?\[\]\:\|\*]','_',name)
            shutil.move(name, newName)
            print('开始转换 ' + name)
            excel = win32.DispatchEx('Excel.Application')
            nameAbsPath = os.path.abspath(newName)
            try:
                wb = excel.Workbooks.Open(nameAbsPath)
                wb.SaveAs(nameAbsPath + "x", FileFormat=51)
                wb.Close()
                excel.Application.Quit()
                print(newName + '转换完成')
            except:
                # print('存在重名文件' + nameAbsPath + "x")
                print(traceback.format_exc())
                # os.unlink(newName)
def changeFile(name):
    print('开始转换 ' + name)
    excel = win32.DispatchEx('Excel.Application')
    nameAbsPath = os.path.abspath(name)
    try:
        wb = excel.Workbooks.Open(nameAbsPath)
        wb.SaveAs(nameAbsPath + "x", FileFormat=51)
        wb.Close()
        excel.Application.Quit()
        print(name + '转换完成')
    except:
        # print('存在重名文件' + nameAbsPath + "x")
        print(traceback.format_exc())

# 将path路径下的所有excel文件(后缀为xls或xlsx的文件)合并成只有一个sheet的excel文件
# path为绝对路径
# headCounts为表头有几行
# tailCounts为表尾有几行
def heBingExcel(path, finalName, headCounts, tailCounts):
    finalName = finalName + '.xlsx'
    # writer = pd.ExcelWriter(finalName)
    print('开始合并：')
    count = 0
    dfTemp = pd.DataFrame()
    for name in os.listdir(path):
        if name.endswith('xlsx'):
            temp = pd.read_excel(name, header=headCounts - 1, skipfooter=tailCounts)
            print('读取 ' + name + '完成')
            dfTemp = pd.concat([dfTemp, temp])
            count = count + 1
    # print(dfTemp)
    print('合并完成，正在生成' + finalName + '...')
    dfTemp.to_excel(finalName, index=None)
    print('已生成' + finalName)


def find1(hebing):
    keys = {'徐州', '睢宁'}
    for key in keys:
        if key in hebing:
            return True


if __name__ == '__main__':
    # 表样式：社会信用代码（纳税人识别号）	纳税人名称	完税证明号码	车辆识别代号	车辆档案编号	发动机号码	牌照号码	车辆类型	购置日期	车辆厂牌型号	机动车销售统一发票不含税价格合计	车辆购置税计税方式	税务机关经办人	车辆购置税免（减）税条件	减免税额	免税填发日期	录入人	主管税务所（科、分局）	受理申报税务机构	征收机关
    # 将本目录下的xls文件全部转换为xlsx文件
    # toXlsx('.')

    # 将本目录下的xlsx文件合并为文件名为“合并.xlsx”的文件，且去除每个数据区表尾2行
    # heBingExcel(绝对目录，合并文件名，表栏目为第几行，表尾有几行)
    # print('')
    headCount = input("输入表栏目所在行数:")
    tailCount = input("输入表尾所在行数:")
    heBingExcel('.', '合并', int(headCount), int(tailCount))
    # heBingExcel('.','合并',3,2)

    # hebing = pd.read_excel('合并.xlsx', thousands=',')
    # print(hebing.head())

    # 统计某些列总计
    # print(hebing['减免税额'])
    # print(hebing['减免税额'].sum())
    # print(hebing['机动车销售统一发票不含税价格合计'].sum())
    # print(hebing['社会信用代码（纳税人识别号）'])

    # 查找特定列包不包含多字符，使用lambda表达式
    # print(hebing[hebing['纳税人名称'].apply(lambda x:x if '徐州' in x else 0)!=0])
    # keys = {'徐州','睢宁'}
    # flag = hebing['纳税人名称'].apply(lambda x: any(key in x for key in keys))
    # print(hebing[flag==True]['纳税人名称'])

    # 查找特定列包不包含多字符，不使用使用lambda表达式
    # flag = hebing['纳税人名称'].apply(find1)
    # print(hebing[flag==True]['纳税人名称'])

    # 多条件筛选：筛选出发票金额大于80w的徐州企业
    # flag1 = hebing['机动车销售统一发票不含税价格合计'].apply(lambda x:True if x > 800000 else False)
    # flag2 = hebing['纳税人名称'].apply(lambda x : True if '徐州' in x else False)
    # print(hebing[(flag1==True)&(flag2==True)]['纳税人名称'])

    # 多表联合查询
    # 查询合并.xlsx纳税人在清册1.xlsx中缴纳的税款
    # qingce1 = pd.read_excel('车辆购置税减免税清册1.xlsx',header=2,skipfooter=2, thousands=',')
    # temp1 = hebing.merge(qingce1,how='outer',on='完税证明号码')
    #
    # # temp1.to_excel('全外连接.xlsx',index=None)
    # print(temp1.columns.values)
    # newDF = pd.DataFrame(temp1,columns=['社会信用代码（纳税人识别号）_x','社会信用代码（纳税人识别号）_y'
    #     ,'纳税人名称_x','纳税人名称_y','完税证明号码','减免税额_x','减免税额_y'])
    # print(newDF.columns)
    # # newDF.to_excel('全外连接结果表.xlsx', index=None)
    # newDF['idFlag'] = newDF[['社会信用代码（纳税人识别号）_x','社会信用代码（纳税人识别号）_y']].apply(lambda x:True if x['社会信用代码（纳税人识别号）_x'] == x['社会信用代码（纳税人识别号）_y'] else False,axis=1)
    # print(newDF['idFlag'])
    # newDF.to_excel('全外连接结果表-对比后.xlsx', index=None)

    # 可视化
    # 生产商分组数量统计
    # hebing['生产商'] = hebing['车辆厂牌型号'].apply(lambda x: re.sub('[A-Za-z0-9\-\/\牌\(\)]', '', str(x)))
    # # print(hebing['生产商'])
    # print(hebing.groupby('生产商').agg('size').sort_values(ascending=False))
    #
    # # 统计生产商为徐工的价格区间，并统计个数
    # xugongDF = pd.DataFrame()
    # xugongDF['徐工价格区间'] = hebing.groupby('生产商').get_group('徐工')['机动车销售统一发票不含税价格合计'].apply(lambda x: x // 100000)
    # print(xugongDF.groupby('徐工价格区间').agg('size').sort_values(ascending=False))
    # xugongDF.groupby('徐工价格区间').agg('size').plot()
    # # plt.show()
    # # 统计整体价格区间并统计个数
    # hebing['价格区间'] = hebing['机动车销售统一发票不含税价格合计'].apply(lambda x: x // 100000)
    # print(hebing.groupby('价格区间').agg('size'))
    # # 画图
    # hebing.groupby('价格区间').agg('size').plot()
    # plt.show()

    os.system('pause')