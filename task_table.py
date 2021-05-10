#!/usr/bin/env python3
# -*- coding: utf-8 -*-
__author__ = 'Zhou Fall'

import os,sys
import pandas as pd    #数据处理
import openpyxl
import numpy as np
import matplotlib.pyplot as plt
import ctypes
import warnings
warnings.filterwarnings("ignore")     #用于排除警告

plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
plt.rcParams['axes.unicode_minus']=False #用来正常显示负号
plt.rcParams["font.size"]=12   #font.size 字体的大小（默认10）
# plt.style.use('fivethirtyeight')
# plt.style.use('ggplot')    #风格还不错
# plt.style.use('seaborn-colorblind')
# plt.style.use('seaborn-bright')
# plt.style.use('seaborn-pastel')
# 设置绘图样式
# 参照下方配色方案，第三参数为颜色数量，这个例子的范围是3-12，每种配色方案参数范围不相同
# bmap = brewer2mpl.get_map('Set3', 'qualitative', 10)
# colors = bmap.mpl_colors
# plt.rcParams['axes.prop_cycle'] = colors
# print(plt.rcParams.keys())
# FontSize=15    #设置字体大小

df_concated_time = []
df_concated_money = []

#获取excel的相对路径，并做简单判断是否为空
def get_excel_path():
    excel_path = ''
    filenames = os.listdir()  # 使用当前默认路径，将导出的xlsx和脚本放一起
    for file in filenames:
        if file.endswith('xlsx'):
            excel_path = file
    if excel_path == '':
        MessageBox = ctypes.windll.user32.MessageBoxW
        MessageBox(0, '当前路径下没有TB下载下来的工时统计表格', '错误提示', 0)
        sys.exit(0)
    return excel_path

#获取excel文件的sheet名，即group list
def get_group_list(excel_path):
    sheet_name = []
    df = pd.read_excel(excel_path, sheet_name=None)  # sheet_name = None， 读所有的sheet
    for k, v in df.items():
        sheet_name.append(k)
    return sheet_name

#初步筛选，将需要的字段保留，并导出到新的excel
def pick_up_data(sheet_name):
    global df_concated_time,df_concated_money
    # 创建一个空的output文件夹
    if not os.path.exists('output'):
        os.makedirs('output')
    else:
        files = os.listdir('output')  # 使用当前默认路径，将导出的xlsx和脚本放一起
        for file in files:
            os.remove('output\\' + file)
    dfs_time = []
    dfs_money = []
    for name in sheet_name:
        # print(name)
        df = pd.read_excel(excel_path,sheet_name=name)    #sheet_name = None， 读所有的sheet
        df = df[df['任务类型'] =='PA-on-site日志表单']    #选取任务类型是onsite日志表单的
        df = df.sort_values(by=['列表','截止时间'])    #按照列表进行排序，列表就是每个人的人名，执行者，参与者，创建者都有可能有邮箱在里面，使用列表来代表人
        #筛选出工时统计感兴趣的数据
        df_time = df[['列表','标题','\"工作内容\"','截止时间','\"正常工时\"','\"请假时长\"','\"加班时长\"','\"加班性质\"','\"加班原因\"']]
        df_time_rename = ['姓名','标题','工作内容','日期','正常工时','请假时长','加班时长','加班性质','加班原因']
        df_time.columns = df_time_rename
        df_time = df_time.copy()
        df_time.loc[:,'正常工时'] = df_time.loc[:,'正常工时'].fillna(0)  #空值补0
        df_time.loc[:, '请假时长'] = df_time.loc[:, '请假时长'].fillna(0)  # 空值补0
        df_time.loc[:, '加班时长'] = df_time.loc[:, '加班时长'].fillna(0)  # 空值补0
        outputpath1 = 'output\技术三科工时填写.xlsx'
        try:
            # 如下操作只是为了不覆盖原有的Excel
            book = openpyxl.load_workbook(outputpath1)
            writer = pd.ExcelWriter(outputpath1, engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df_time.to_excel(writer, index=False, sheet_name=name)  # index =False 去掉索引号
            writer.save()
        except:
            # 直接写入方式，会覆盖原有的excel，适用于创建
            df_time.to_excel(outputpath1, index=False, sheet_name=name)  # index =False 去掉索引号

        # 正常工时，加班时长按照人名汇总,做简单统计
        df_time_temp = df_time.groupby('姓名', as_index=False)[['正常工时', '请假时长', '加班时长']].sum()
        # print(df_time_temp)
        df_time_temp['加班占比'] = (df_time_temp['加班时长'] / df_time_temp['正常工时'] ).round(3) * 100
        df_time_temp['加班占比'] = df_time_temp['加班占比'].fillna(0)  # 空值换成0
        df_time_temp['加班占比'][df_time_temp['加班占比'] < 0] = 0  # 强啊，使用这种方法定位，df_time_temp['加班占比'] < 0是一串true,false!!!
        df_time_temp['分组'] = name
        dfs_time.append(df_time_temp)
        #筛选出费用统计感兴趣的数据
        df_money = df[['列表','标题','截止时间','\"加班交通费\"','\"工作（出差）地点\"','\"出差餐费（有发票）\"','\"出差餐补（无发票）\"','\"出差交通费\"','\"出差住宿费\"','\"其他项目相关费用\"']]
        df_money_rename = ['姓名','标题','日期','加班交通费','工作（出差）地点','出差餐费（有发票）','出差餐补（无发票）','出差交通费','出差住宿费','其他费用']
        df_money.columns = df_money_rename
        outputpath2 = 'output\技术三科费用填写.xlsx'
        try:
            book = openpyxl.load_workbook(outputpath2)
            writer = pd.ExcelWriter(outputpath2, engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df_money.to_excel(writer, index=False, sheet_name=name)  # index =False 去掉索引号
            writer.save()
        except:
            # 直接写入方式，会覆盖原有的excel，适用于创建
            df_money.to_excel(outputpath2, index=False, sheet_name=name)  # index =False 去掉索引号

        # 删掉费用几栏里全部为空的数据,Todo:填写费用的逻辑还需要完善
        df_money = df_money.copy()    #为了log不输出warning，可不加
        df_money[['出差餐费（有发票）','出差餐补（无发票）', '出差交通费', '出差住宿费', '加班交通费','其他费用']] = df_money[['出差餐费（有发票）', '出差餐补（无发票）','出差交通费', '出差住宿费', '加班交通费','其他费用']].replace(0,None)
        df_money_temp = df_money.dropna(axis='index', how='all', subset=['出差餐费（有发票）','出差餐补（无发票）', '出差交通费', '出差住宿费', '加班交通费','其他费用'])
        dfs_money.append(df_money_temp)

    #按行合并
    df_concated_time = pd.concat(dfs_time)
    df_concated_money = pd.concat(dfs_money)
    outputpath3 = 'output\工时统计结果.xlsx'
    df_concated_time.to_excel(outputpath3, index=False, sheet_name='工时统计')  # index =False 去掉索引号
    outputpath4 = 'output\费用统计结果.xlsx'
    df_concated_money.to_excel(outputpath4, index=False, sheet_name='费用统计')  # index =False 去掉索引号

def check_TB_data():
    df = pd.read_excel('output\技术三科工时填写.xlsx',sheet_name=None)
    keys = list(df.keys())
    df_concat = pd.DataFrame()
    for i in keys:
        df1 = df[i]
        df_concat = pd.concat([df_concat, df1])
    df_concat.sort_values(by=['姓名'])
    # print(df_concat)
    # 错误一，非休息日加班，正常工时+请假时长 ！=8
    # 错误二，加班时长=0，加班性质不是无加班
    # 错误三，休息日加班，正常工时不等于加班时长

    df_error1 = df_concat.loc[(df_concat['加班性质'] != '休息日加班') & (df_concat['正常工时'] + df_concat['请假时长'] !=8) ,:]
    df_error2 = df_concat.loc[(df_concat['加班性质'] != '无加班') & (df_concat['加班时长'] ==0) ,:]
    df_error3 = df_concat.loc[(df_concat['加班性质'] == '休息日加班') & (df_concat['正常工时']!=0) ,:]

    df_concated_error = pd.concat([df_error1,df_error2,df_error3])
    df_concated_error = df_concated_error.drop_duplicates()
    outputpath = 'output\工时填写错误.xlsx'
    df_concated_error.to_excel(outputpath, index=False, sheet_name='工时填写错误')  # index =False 去掉索引号

def autolabel(rects,ax):
    # attach some text labels
    for rect in rects:
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width()/2.0, 1.03*height,'%.1f' % float(height),ha='center', va='bottom')

def autolabelh(rects,ax):
    # attach some text labels
    for rect in rects:
        width = rect.get_width()
        # height = rect.get_height()
        # print(rect.get_y())
        ax.text(1.03*width,rect.get_y(), '%d' % int(width),ha='center', va='bottom')

def get_color(x, y):
    color = []
    for i in range(len(x)):
        if y[i] < 25:
            color.append("green")
        elif y[i] < 33:
            color.append("lightseagreen")
        elif y[i] < 50:
            color.append("gold")
        else:
            color.append("coral")
    return color

def draw_picture(df_time,group):
    plt.style.use('seaborn-bright')
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2,figsize=(14, 8))
    fig.subplots_adjust(wspace=0.18, hspace=0.3,left=0.1, right=0.9,top=0.93, bottom=0.1)  #调整缩放比例，右侧图标可以完全显示出来，right=0.8默认0.9
    number = len(group)
    ax = [ax1,ax2,ax3,ax4]
    for i in range(number):
        df_temp = df_time.loc[df_time['分组'] == group[i]]
        x_label = df_temp['姓名']
        # print(x_label)
        x = np.arange(len(x_label))
        y1 = df_temp['正常工时']
        y2 = df_temp['加班时长']
        total_width, n = 0.8, 2
        width = total_width / n
        figure1 = ax[i].bar(x, y1, width=width, label='正常工时')  # bar是垂直方向的，barh是水平方向的
        figure2 = ax[i].bar(x + width, y2, width=width, label='加班时长')
        autolabel(figure1, ax[i])
        autolabel(figure2, ax[i])
        ax[i].set_xticks(x + width/2)
        ax[i].set_xticklabels(x_label,rotation=50)    #单图形时，还可以plt.xticks(rotation=50)
        ax[i].set_ylim([0, 260])
        ax[i].set(ylabel='时间 (小时)', title=group[i])
        # ax[i].grid(color='r', linestyle='-', linewidth=2)    #设置网格

        # y3 = df_time['加班占比']
        #color='g' , 默认颜色是blue，什么都不加的时候也会把下面这两个图区分开
        # figure3 = ax.bar(x+width*2, y3,width=width,label='加班占比')
        # autolabel(figure3,ax)
        # plt.tight_layout()

    # 设置图例在图形外面
    ax[1].legend(loc=2, bbox_to_anchor=(1.03, 1.0), borderaxespad=0, numpoints=1, fontsize=10)
    plt.savefig('output\技术三科工时统计图',dpi=600,bbox_inches="tight")    #设置保存的图片大小

    #画另一张加班占比图
    plt.style.use('ggplot')    #风格还不错
    fig,ax5 = plt.subplots(figsize=(14, 8))
    df_time = df_time.sort_values(by=['加班占比'])
    x5 = df_time['姓名']
    y5 = df_time['加班占比']
    y_pos = np.arange(len(x5))
    figure5 = ax5.barh(y_pos, y5,color=get_color(x5.values,y5.values) ,align='center')
    ax5.set_yticks(y_pos)
    ax5.set_yticklabels(x5)
    ax5.set_xlabel('百分比%')
    # ax5.set_ylabel('姓名')
    autolabelh(figure5,ax5)
    ax5.set_title('加班占比统计图')
    plt.savefig('output\技术三科加班统计图', dpi=600, bbox_inches="tight")  # 设置保存的图片大小
    # plt.ion()
    # plt.pause(10)  # 显示秒数
    # plt.close()
    plt.show()

if __name__ == '__main__':
    excel_path = get_excel_path()
    group_list = get_group_list(excel_path)
    print(group_list)
    #选取感兴趣需要的数据
    pick_up_data(group_list)
    # 对数据内容进行校验
    check_TB_data()
    #将结果可视化输出
    # draw_picture(df_concated_time,group_list[0:4])    #画工时图和加班占比图


