#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:PC
# datetime:2022/9/12 7:00
# software: PyCharm

import pandas as pd
# import numpy as np
# import configparser
import os
# import shutil
import glob
import logging
import datetime
import traceback
logging.basicConfig(filename='./log.log', filemode='w')
try:

    # cp = configparser.ConfigParser()
    # cp.read('./config.cfg', encoding='utf-8-sig')
    # 写入淘汰产品汇总表
    # df_useless = pd.read_excel('./淘汰MSKU汇总.xlsx')
    # df_useless['备货天数淘汰'] = 0
    # 写入关键字表
    # df_keys = pd.read_excel('./关键字对应表.xlsx')
    # print(df_keys.dtypes)
    # 写入MSKU备货天数
    df_msku_day = pd.read_excel('./MSKU备货天数汇总.xlsx')

    df_name = pd.read_excel('./MSKU负责人.xlsx')
    # 遍历文件夹得到文件名
    data_path = './asinking数据文件夹'
    data_list = glob.glob(os.path.join(data_path, "*.xlsx"))
    print(data_list)

    dl = []
    for d in data_list:
        dl.append(pd.read_excel(d))
    # print(dl)
    data = pd.concat(dl).reset_index(drop = True)
    data['欧洲/北美汇总行'].fillna('', inplace = True)
    title_list = ['欧洲/北美汇总行',
                  'MSKU',
                  '店铺',
                  '国家',
                  '品名',
                  'FBA库存',
                  'FBA在途',
                  '7天销量',
                  '7天日均']
    # 清洗数据为0行
    data = data.drop(data[(data['FBA库存'] == 0) & (data['FBA在途'] == 0) & (data['30天销量'] == 0)].index)
    data = data.reset_index(drop=True)
    data = data[title_list]
    # print(data.dtypes)
    # 把MSKU和品名列取出做一个新的一对一映射关系，再通过表连接对应新的品名关系
    msku_name = data[['MSKU', '品名']].dropna()
    msku_name = msku_name.drop_duplicates(subset = ['MSKU'], keep = 'first')
    msku_name = msku_name.reset_index(drop = True)
    # 全行对应品名关系
    data.drop(columns=['品名'], inplace=True)
    # print(data)
    data = pd.merge(data, msku_name, how = 'left', on = 'MSKU')
    # 清洗汇总行数据
    if '共享库存' in data['欧洲/北美汇总行'].values:
        data = data.drop(data[(data['欧洲/北美汇总行'] == '共享库存')].index)
        data = data.reset_index(drop = True)
        # eu_store = data['店铺'].iloc[data[(data['欧洲汇总行'] == '是')].index].apply(lambda x: x.split('-')[0]).tolist()
        data.loc[data['欧洲/北美汇总行'] == '是', '店铺'] = data['店铺'].iloc[data[(data['欧洲/北美汇总行'] == '是')].index].apply(lambda x: x.split('-')[0])
    # 清洗美国站数据，把所有单独上架在美国站的US-US改为US，其他不处理
    # if '美国' in data['国家'].values:
    #     data = data.drop(data[(data['国家'] == '墨西哥') | (data['国家'] == '加拿大')].index)
    #     data = data.reset_index(drop = True)
    data.loc[(data['国家'] == '美国'), '店铺'] = \
        data['店铺'].iloc[data[(data['国家'] == '美国')].index].apply(lambda x: x.split('-')[0])
    # 清洗澳洲站，数据；更新时间2022-09-12
    data.loc[(data['国家'] == '澳洲'), '店铺'] = \
        data['店铺'].iloc[data[(data['国家'] == '澳洲')].index].apply(lambda x: x.split('-')[0])
    # 用7天销量计算7天日均，并替换到7天日均
    data['7天日均'] = round(data['7天销量']/7, 2)
    # print(data)
    # 新建一列需求数量
    data["备货数量"] = ''
    # data['备货天数'] = ''
    # data['是否按统一备货天数'] = ''
    data['行号'] = ''
    data['I'] = 'I'
    data['H'] = 'H'
    data['F'] = 'F'
    data['G'] = 'G'
    data['备货原因'] = ''
    data['备货天数 '] = ''
    data['日均'] = ''
    data['备货数量 '] = ''

    # df_keys_list = df_keys['包含关键字'].tolist()
    # df_useless_list = df_useless['MSKU'].tolist()
    # print(df_keys_list)
    # count = 0
    # 写入配置文件权重
    # high_weight = int(cp.get('Weight', 'high'))
    # low_weight = int(cp.get('Weight', 'low'))
    # for k in df_keys_list:
    #     data.loc[(data['品名'].str.contains(k, case = False, regex = False)) &
    #              (data['7天日均'] >= high_weight), '备货天数'] = df_keys['备货天数(高权重)'][count]
    #     data.loc[(data['品名'].str.contains(k, case = False, regex = False)) &
    #              ((data['7天日均'] < high_weight) & (data['7天日均'] >= low_weight)), '备货天数'] = df_keys['备货天数(中权重)'][count]
    #     data.loc[(data['品名'].str.contains(k, case = False, regex = False)) &
    #              (data['7天日均'] < low_weight), '备货天数'] = df_keys['备货天数(低权重)'][count]
    #     count = count + 1
    # 淘汰产品
    # for l in df_useless_list:
    #     data.loc[(data['MSKU'] == l), '备货天数'] = 0
    # print(data)
    # df_useless = df_useless.drop_duplicates(subset=['店铺', 'MSKU'], keep='first')
    # # print(df_useless)
    # data['店铺'] = data['店铺'].apply(str)
    # data['MSKU'] = data['MSKU'].apply(str)
    # df_useless['店铺'] = df_useless['店铺'].apply(str)
    # df_useless['MSKU'] = df_useless['MSKU'].apply(str)
    df_msku_day = df_msku_day.drop_duplicates(subset=['店铺', 'MSKU'], keep='first')
    df_msku_day = df_msku_day.reset_index(drop=True)
    df_msku_day['店铺'] = df_msku_day['店铺'].str.upper()
    data = pd.merge(data, df_msku_day, how='left', on=['店铺', 'MSKU'])
    # data.loc[(data['备货天数淘汰'] == 0), '备货天数'] = 0
    # print(data)
    data['备货天数'].fillna(0, inplace=True)
    # data.loc[(data['备货天数'] == ''), '是否按统一备货天数'] = '是'
    # data.loc[(data['备货天数'] == '') & (data['7天日均'] >= high_weight), '备货天数'] = int(cp.get('SupplyDays', 'GTF'))
    # data.loc[(data['备货天数'] == '') & ((data['7天日均'] < high_weight) & (data['7天日均'] >= low_weight)), '备货天数'] \
    #     = int(cp.get('SupplyDays', 'TTF'))
    # data.loc[(data['备货天数'] == '') & (data['7天日均'] < low_weight), '备货天数'] = int(cp.get('SupplyDays', 'LTT'))
    # data.loc[(data['是否按统一备货天数'] == ''), '是否按统一备货天数'] = '否'
    # 以上步骤数据清洗完毕，下面为分配数据
    # =I2 * H2 - F2 - G2
    # 输出品名为空的文件
    data_ud = data[data['品名'].isnull()]
    # data_ud = data_ud.drop(data_ud[(data_ud['备货天数淘汰'] == 0)].index)
    data_ud = data_ud[['店铺', 'MSKU', '品名']]
    data_ud.to_excel('./品名缺失.xlsx', index=False)
    data = data.dropna(subset = ['品名'])
    final_title = ['负责人',
                   '店铺',
                   'MSKU',
                   '品名',
                   '备货数量',
                   'FBA库存',
                   'FBA在途',
                   '7天日均',
                   '备货天数',
                   '备货原因',
                   '备货天数 ',
                   '日均',
                   '备货数量 ']

    df_name = df_name.drop_duplicates(subset=['MSKU'], keep='first')
    data['MSKU'] = data['MSKU'].apply(str)
    df_name['MSKU'] = df_name['MSKU'].apply(str)
    data = pd.merge(data, df_name, how='left', on='MSKU')
    data = data[final_title]
    data['负责人'].fillna('无', inplace=True)
    person_list = data['负责人'].unique()

    data = data.drop_duplicates(subset=['店铺', 'MSKU'], keep='first')
    data = data.sort_values(by = ['店铺', 'MSKU'], ascending = False)
    data = data.reset_index(drop=True)
    data["行号"] = data.index + 2
    data["备货数量"] = '=' + data["备货天数"].apply(str) + r'*H' + data["行号"].apply(str) + r'-F' + data["行号"].apply(
        str) + r'-G' + data["行号"].apply(str)
    # data["备货数量 "] = '=' + data["备货天数 "].apply(str) + r'*L' + data["行号"].apply(str) + r'-F' + data["行号"].apply(
    #     str) + r'-G' + data["行号"].apply(str)
    # data["备货数量"] = data['备货天数'] * data['7天日均'] - data['FBA库存'] - data['FBA在途']

    # 若有执行权限问题用管理员方式运行，360报毒，不拿权限了
    # def readonly_handler(func, path):
    #     os.chmod(path, stat.S_IWRITE)
    #     func(path)
    data = data[final_title]

    def create_dir_not_exist(path):
        if not os.path.exists(path):
            # shutil.rmtree(path, onerror = readonly_handler)
            os.mkdir(path)


    # print(data.dtypes)
    # print(data)
    # 总表
    todaydate = datetime.datetime.now().strftime('%Y%m%d')
    name_all = '汇总' + todaydate
    data.to_excel('./' + name_all + '.xlsx', index=False)
    # 分表
    create_dir_not_exist("./所有负责人" + todaydate)
    for s in person_list:
        data_temp = data.loc[data['负责人'] == s]
        data_temp = data_temp.reset_index(drop=True)
        data_temp["行号"] = data_temp.index + 2
        data_temp["备货数量"] = '=' + data_temp["备货天数"].apply(str) + r'*H' + data_temp["行号"].apply(
            str) + r'-F' + data_temp["行号"].apply(str) + r'-G' + data_temp["行号"].apply(str)
        data_temp = data_temp[final_title]
        data_temp.to_excel('./所有负责人' + todaydate + r'/{}'.format(s) + todaydate + '.xlsx', index=False)

    # print(person_list)
    print('success')

except:
    s = traceback.format_exc()
    logging.error(s)
