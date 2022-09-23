#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:PC
# datetime:2021/3/24 7:00
# software: PyCharm

import pandas as pd
import numpy as np
import os
import logging
import traceback
logging.basicConfig(filename='./log.log', filemode='w')
try:
    # 提示是否有信用卡,0否1是
    flag = 0
    in_num = input('广告费是否信用卡扣款,输入数字1为是,输入其他默认为否:')

    if str(1) == in_num:
        flag = 1
    # print(flag)
    # 写入汇率
    df_exchange = pd.read_excel('./汇率.xlsx')
    df_exchange = df_exchange.drop_duplicates(subset=['币种'], keep='first')
    print(df_exchange.info())
    # 写入成员架构表，并生成相应列表
    df_frame = pd.read_excel('./成员架构表.xlsx')

    # 写入国家站点对应
    # df_site = pd.read_excel('./国家站点对应.xlsx')
    # df_site = df_site.drop_duplicates(subset=['国家'], keep='first')
    data = pd.read_excel('./原始数据.xlsx', header = 1, sheet_name = "sheet1")
    data_all = pd.read_excel('./原始数据.xlsx', header=1, sheet_name="sheet1")
    # data = data.reset_index(drop = True)
    # data.to_excel('./' + data["月份"][0] + '1利润报表.xlsx', index=False)
    # 判断月份
    data["month"] = data["月份"].apply(lambda x: x.split('-')[1])
    data["year"] = data["月份"].apply(lambda x: x.split('-')[0])
    month = data["month"][0]
    year = data["year"][0]

    if int(year) % 4 == 0 and int(year) % 400 == 0:
        print('闰年')
        df_month = pd.DataFrame(data=[['01', 31],
                                      ['02', 29],
                                      ['03', 31],
                                      ['04', 30],
                                      ['05', 31],
                                      ['06', 30],
                                      ['07', 31],
                                      ['08', 31],
                                      ['09', 30],
                                      ['10', 31],
                                      ['11', 30],
                                      ['12', 31]],
                                columns=['month', '这个月天数'])
    elif int(year) % 4 == 0 and int(year) % 100 != 0:
        print('闰年')
        df_month = pd.DataFrame(data=[['01', 31],
                                      ['02', 29],
                                      ['03', 31],
                                      ['04', 30],
                                      ['05', 31],
                                      ['06', 30],
                                      ['07', 31],
                                      ['08', 31],
                                      ['09', 30],
                                      ['10', 31],
                                      ['11', 30],
                                      ['12', 31]],
                                columns=['month', '这个月天数'])
    else:
        print('平年')
        df_month = pd.DataFrame(data=[['01', 31],
                                      ['02', 28],
                                      ['03', 31],
                                      ['04', 30],
                                      ['05', 31],
                                      ['06', 30],
                                      ['07', 31],
                                      ['08', 31],
                                      ['09', 30],
                                      ['10', 31],
                                      ['11', 30],
                                      ['12', 31]],
                                columns=['month', '这个月天数'])

    # 百分比函数
    def turn_per(x):
        return '%.2f%%' % (x * 100)

    # 国家对应站点
    df_site = pd.DataFrame(data=[['美国', '北美站'],
                                 ['英国', '欧洲站'],
                                 ['德国', '欧洲站'],
                                 ['意大利', '欧洲站'],
                                 ['西班牙', '欧洲站'],
                                 ['法国', '欧洲站'],
                                 ['日本', '日本站'],
                                 ['墨西哥', '北美站'],
                                 ['加拿大', '北美站'],
                                 ['荷兰', '欧洲站'],
                                 ['瑞典', '欧洲站'],
                                 ['波兰', '欧洲站']],
                           columns=['国家', '站点'])
    # 取所有负责人列表
    # print(data["负责人"])
    data["负责人"] = data["负责人"].fillna(';')
    data["负责人"] = data["负责人"].apply(lambda x: x.split(';')[0])
    data_all["负责人"] = data["负责人"].apply(lambda x: x.split(';')[0])

    data = pd.merge(data, df_month, how="left", on="month")
    data = pd.merge(data, df_exchange, how="left", on="币种")
    data = pd.merge(data, df_site, how="left", on="国家")
    data = pd.merge(data, df_frame, how="left", on="负责人")
    data['经理/主管'].fillna("无", inplace=True)
    data['组长'].fillna("无", inplace=True)
    data['负责人'].fillna("无", inplace=True)
    data_all = pd.merge(data_all, df_frame, how="left", on="负责人")

    frame_list_a = data['经理/主管'].unique()
    # frame_list_c = df_frame['负责人'].unique()
    print(data)
    data["日均销量"] = data["FBA销量"]/data["这个月天数"]
    data.loc[data['站点'] == '欧洲站', '销售额（RMB）'] = (data["退款"] + data["FBA销售额"] + data["销售税-商品销售税"]) * data["汇率"]
    data.loc[(data['站点'] == '北美站') | (data['站点'] == '日本站'), '销售额（RMB）'] = (data["退款"] + data["FBA销售额"]) * data["汇率"]
    # data["销售额（RMB）"] = (data["退款"] + data["FBA销售额"]) * data["汇率"]
    data["毛利润（RMB）"] = data["毛利润"] * data["调整汇率"]
    data["广告费（本币）"] = data["SP广告费"] + data["SD广告费"] + data["SB广告费分摊"] + data["SBV广告费分摊"] + data['广告费-调整分摊']
    data["日均广告费"] = data["广告费（本币）"] / data["这个月天数"]
    data["Coupon+秒杀费（本币）"] = data["优惠券"] + data["秒杀费"] + data["早期评论者计划"]
    # data["损耗费RMB"] = data["损耗费"] * data["调整汇率"]
    data["采购成本"] = (data["库存调整-初始化调整"] +
                    data["采购成本-初始化调整"] +
                    data["FBA订单发货-采购成本"] +
                    data["FBM订单发货-采购成本"] +
                    data["买家退货-采购成本"] +
                    data["库存调整-采购成本"] +
                    data["库存移除-采购成本"] +
                    data["库存差异-采购成本"] +
                    data["FBA补货差异-采购成本"]) * data["调整汇率"]

    data["物流+仓库成本"] = (data["FBM售出商品-头程费用"] +
                       data["FBA售出商品-头程费用"] +
                       data["买家退货-头程费用"] +
                       data["库存调整-头程费用"] +
                       data["库存移除-头程费用"] +
                       data["库存差异-头程费用"] +
                       data["FBA补货差异-头程费用"] +
                       data["海外仓-其他费用"]) * data["调整汇率"]

    data["毛利率"] = (data["毛利润（RMB）"] / data["销售额（RMB）"]).apply(turn_per)
    data.loc[data['销售额（RMB）'] == 0, '毛利率'] = '0%'

    # data["退货率"] = ((data["退货可售量"] + data["退货不可售量"]) / (data["FBA销量"] + data["库存赔偿量"])).apply(turn_per)
    # data["退款率"] = (data["退款"] / (data["FBA销售额"] + data["MSKU赔偿"])).apply(turn_per)
    if 1 == flag:
        data["回款额（RMB）"] = (data["FBA销售额"] +
                            data["FBA买家运费"] +
                            data["促销折扣"] +
                            data["MSKU赔偿"] +
                            data["费用分摊"] +
                            data["退款"] +
                            data["售出订单-平台费"] +
                            data["售出订单-FBA发货费"] +
                            data["售出订单-FBM发货费"] +
                            data["售出订单-其他费"] +
                            data["售出订单-运输标签费"] +
                            data["多渠道订单-FBA发货费"] +
                            data["退款订单-平台费"] +
                            data["退款订单-发货费"] +
                            data["退款订单-其他费"] +
                            data["SP广告费"] +
                            data["SD广告费"] +
                            data["SB广告费分摊"] +
                            data["SBV广告费分摊"] +
                            data["广告费差异分摊"] +
                            data["优惠券"] +
                            data["秒杀费"] +
                            data["早期评论者计划"] +
                            data["月仓储费"] +
                            data["长期仓储费"] +
                            data["差异分摊"] +
                            data["超量仓储费"] +
                            data["合作承运费"] +
                            data["销毁费"] +
                            data["移除费"] +
                            data["退货入仓费"] +
                            data["FBA仓储差异分摊"] +
                            data["库存调整费"] +
                            data["库存调整差异分摊"] +
                            data["FBA物流国际货运费"] +
                            data["合仓费"] +
                            data["订阅费"] +
                            data["平台其他费分摊"] +
                            data["销售税-隐藏税"] +
                            data["销售税-商品销售税"] +
                            data["销售税-运费税"] +
                            data["销售税-礼品包装税"] +
                            data["销售税-促销折扣税"] +
                            data["亚马逊代收代扣-市场税"] +
                            data["亚马逊代收代扣-LVIG"] +
                            data["预扣所得税"] +
                            data["TCS_CGST"] +
                            data["TCS_SGST"] +
                            data["TCS_IGST"] +
                            data["测评费"] +
                            data["广告费-信用卡扣款"]) * data["调整汇率"]
    else:
        data["回款额（RMB）"] = (data["FBA销售额"] +
                            data["FBA买家运费"] +
                            data["促销折扣"] +
                            data["MSKU赔偿"] +
                            data["费用分摊"] +
                            data["退款"] +
                            data["售出订单-平台费"] +
                            data["售出订单-FBA发货费"] +
                            data["售出订单-FBM发货费"] +
                            data["售出订单-其他费"] +
                            data["售出订单-运输标签费"] +
                            data["多渠道订单-FBA发货费"] +
                            data["退款订单-平台费"] +
                            data["退款订单-发货费"] +
                            data["退款订单-其他费"] +
                            data["SP广告费"] +
                            data["SD广告费"] +
                            data["SB广告费分摊"] +
                            data["SBV广告费分摊"] +
                            data["广告费差异分摊"] +
                            data["优惠券"] +
                            data["秒杀费"] +
                            data["早期评论者计划"] +
                            data["月仓储费"] +
                            data["长期仓储费"] +
                            data["差异分摊"] +
                            data["超量仓储费"] +
                            data["合作承运费"] +
                            data["销毁费"] +
                            data["移除费"] +
                            data["退货入仓费"] +
                            data["FBA仓储差异分摊"] +
                            data["库存调整费"] +
                            data["库存调整差异分摊"] +
                            data["FBA物流国际货运费"] +
                            data["合仓费"] +
                            data["订阅费"] +
                            data["平台其他费分摊"] +
                            data["销售税-隐藏税"] +
                            data["销售税-商品销售税"] +
                            data["销售税-运费税"] +
                            data["销售税-礼品包装税"] +
                            data["销售税-促销折扣税"] +
                            data["亚马逊代收代扣-市场税"] +
                            data["亚马逊代收代扣-LVIG"] +
                            data["预扣所得税"] +
                            data["TCS_CGST"] +
                            data["TCS_SGST"] +
                            data["TCS_IGST"] +
                            data["测评费"]) * data["调整汇率"]

    # 保留1位小数
    data['日均销量'] = data['日均销量'].round(1)
    data['销售额（RMB）'] = data['销售额（RMB）'].round(1)
    data['回款额（RMB）'] = data['回款额（RMB）'].round(1)
    data['毛利润（RMB）'] = data['毛利润（RMB）'].round(1)
    data['广告费（本币）'] = data['广告费（本币）'].round(1)
    data['日均广告费'] = data['日均广告费'].round(1)
    data['Coupon+秒杀费（本币）'] = data['Coupon+秒杀费（本币）'].round(1)
    # data['损耗费RMB'] = data['损耗费RMB'].round(1)
    data['采购成本'] = data['采购成本'].round(1)
    data['物流+仓库成本'] = data['物流+仓库成本'].round(1)
    title_list_f = ['月份',
                  'ParentAsin',
                  '店铺',
                  '站点',
                  '品名',
                  '负责人',
                  '分类',
                  '品牌',
                  'FBA销量',
                  '日均销量',
                  '销售额（RMB）',
                  '回款额（RMB）',
                  '广告费（本币）',
                  '日均广告费',
                  'Coupon+秒杀费（本币）',
                  # '损耗费RMB',
                  '采购成本',
                  '物流+仓库成本',
                  # '退货率',
                  '退款率',
                  '毛利率',
                  '毛利润（RMB）']
    title_list = ['月份',
                  'ParentAsin',
                  '店铺',
                  '站点',
                  '品名',
                  '经理/主管',
                  '组长',
                  '负责人',
                  '分类',
                  '品牌',
                  'FBA销量',
                  '日均销量',
                  '销售额（RMB）',
                  '回款额（RMB）',
                  '广告费（本币）',
                  '日均广告费',
                  'Coupon+秒杀费（本币）',
                  # '损耗费RMB',
                  '采购成本',
                  '物流+仓库成本',
                  # '退货率',
                  '退款率',
                  '毛利率',
                  '毛利润（RMB）']

    data = data[title_list]
    data = data.sort_values(by = ['经理/主管', '组长', '负责人', '品名', '店铺'])
    # print(data.dtypes)
    data.to_excel('./' + data["月份"][0] + '利润报表.xlsx', index=False)

    # 为分表判断是否已存在文件夹


    def create_dir_not_exist(path):
        if not os.path.exists(path):
            os.mkdir(path)
    # 分表
    create_dir_not_exist("./所有负责人" + data["月份"][0])
    for p in frame_list_a:
        data_temp = data.loc[data['经理/主管'] == p]
        data_all_temp = data_all.loc[data_all['经理/主管'] == p]
        frame_list_b = data_temp['组长'].unique()
        create_dir_not_exist("./所有负责人" + data["月份"][0] + r'/{}'.format(p))
        for q in frame_list_b:
            data_temp_a = data_temp.loc[data_temp['组长'] == q]
            data_all_temp_a = data_all_temp.loc[data_all_temp['组长'] == q]
            person_list = data_temp_a['负责人'].unique()
            create_dir_not_exist("./所有负责人" + data["月份"][0] + r'/{}'.format(p) + r'/{}'.format(q))
            for s in person_list:
                # sheet12 = pd.ExcelWriter('./所有负责人' + data["月份"][0] + r'/{}'.format(s) + '-利润报表' + data["月份"][0] + '.xlsx')
                data_temp_b = data_temp_a.loc[data_temp_a['负责人'] == s]
                data_temp_b = data_temp_b.reset_index(drop=True)
                data_temp_b = data_temp_b[title_list_f]
                data_all_temp_b = data_all_temp_a.loc[data_all_temp_a['负责人'] == s]
                data_all_temp_b = data_all_temp_b.reset_index(drop=True)
                data_all_temp_b = data_all_temp_b.drop(columns=["经理/主管", "组长"])
                # data_temp.to_excel('./所有负责人' + data["月份"][0] + r'/{}'.format(s) + data["月份"][0] + '.xlsx', index=False)
                data_temp_b.to_excel('./所有负责人' + data["月份"][0] + r'/{}'.format(p) + r'/{}'.format(q) + r'/{}'.format(s)
                                   + '-利润报表' + data["月份"][0] + '.xlsx', index=False)
                # sheet12.save()
                data_all_temp_b.to_excel('./所有负责人' + data["月份"][0] + r'/{}'.format(p) + r'/{}'.format(q)
                                       + r'/{}'.format(s) + '-利润报表原始数据' + data["月份"][0] + '.xlsx', index=False)
                # sheet12.save()
    print('success')

except:
    s = traceback.format_exc()
    logging.error(s)
