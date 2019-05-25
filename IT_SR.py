# ！/usr/bin/python3
# -*- coding: utf-8 -*-#
#  @Time    :2019/5/23  
#  @Author  :jun
#  @Email   :1044622996@qq.com      
#  @File    : IT_SR.py
#  @Software: PyCharm Community Edition
#  @Funtion :


# _*_coding:utf-8_*_

import pymssql
import xlwt
from datetime import datetime
# from common.file_path import file_path
import  os
import time

file_time = time.strftime('%Y%m%d')
dir_path = os.path.dirname(os.path.dirname(__file__))

class py_mssql:
    # 1. 建立连接
    def __init__(self):
        host = '192.168.1.13'
        user = 'tianhong'
        password = '!QAZ2wsx'
        self.pymssql = pymssql.connect(host=host, user=user, password=password, database=None, port=1433,
                                       charset='utf8')  # 实例变量

        self.cursor = self.pymssql.cursor()  # 默认以列表的形式显示

    def fetch_all(self, sql):
        self.cursor.execute(sql)  # 4. 执行sql
        result = self.cursor.fetchall()  # 5. 查看结果
        return result

    def close(self):
        self.cursor.close()  # 6. 关闭页面
        self.pymssql.close()  # 7. 关闭数据库

    def file_name(self):
        path = os.path.join(dir_path, 'datas', 'IT需求明细统计_' +  file_time + '.xls')
        return path

def write_data_to_excel(name, sql):
    mssql = py_mssql()
    # 将sql作为参数传递调用get_data并将结果赋值给result,(result为一个嵌套元组)
    result = mssql.fetch_all(sql)
    header = (u'需求类型', u'文档编号', u'IT部门', u'需求提出部门', u'需求提出人员', u'IT部门经理', u'当前节点', u'需求标题', u'提出时间', u'状态', u'结束时间')
    result.insert(0,header)
    # 实例化一个Workbook()对象(即excel文件)
    wbk = xlwt.Workbook()
    # 新建一个名为Sheet1的excel sheet。此处的cell_overwrite_ok =True是为了能对同一个单元格重复操作。
    sheet = wbk.add_sheet('Sheet1', cell_overwrite_ok=True)

    # 遍历result中的没个元素。


    #  设置表格的宽度
    sheet.col(0).width = (20 * 240)  # 需求类型
    sheet.col(1).width = (20 * 300)  # 文档编号
    sheet.col(2).width = (20 * 100)  # IT部门
    sheet.col(3).width = (20 * 180)  # 需求提出部门
    sheet.col(4).width = (20 * 210)  # 需求提出人员
    sheet.col(5).width = (20 * 180)  # IT部门经理
    sheet.col(6).width = (20 * 400)  # 当前节点
    sheet.col(7).width = (20 * 730)  # 需求标题
    sheet.col(8).width = (20 * 280)  # 提出时间
    sheet.col(9).width = (20 * 160)  # 状态
    sheet.col(10).width = (20 * 280)  # 结束时间

    borders = xlwt.Borders()  # Create borders

    borders.left = xlwt.Borders.THIN  # 添加边框-虚线边框
    borders.right = xlwt.Borders.THIN  # 添加边框-虚线边框
    borders.top = xlwt.Borders.THIN  # 添加边框-虚线边框
    borders.bottom = xlwt.Borders.THIN  # 添加边框-虚线边框

    borders.left_colour = 0x40 # 边框上色
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40

    style = xlwt.XFStyle()  # Create style
    style.borders = borders  # Add borders to style

    font0 = xlwt.Font()
    font0.name = '等线'
    font0.bold = True
    font0.height = 220

    font1 = xlwt.Font()
    font1.name = '等线'
    font1.bold = False
    font1.height = 220  # 11号字体   220 / 20

    for i in range(len(result)):
        for j in range(len(result[i])):
            if i ==0:
                each_header = header[j]
                #font = Font('等线', bold=True)  # 字体，等线， bold 加粗
                style.font = font0
                sheet.write(0, j, each_header,style)
            # 将每一行的每个元素按行号i,列号j,写入到excel中
            else:
                if j in (8,10):
                    style.font = font1
                    if result[i][j] != None:
                        a = result[i][j]
                        datestr = datetime.strftime(a,'%Y-%m-%d %H:%M:%S')
                        sheet.write(i, j, datestr,style)
                    else:
                        style.font = font1
                        sheet.write(i, j, 'NUll',style)
                else:
                    style.font = font1
                    sheet.write(i, j, result[i][j],style)

    # 以传递的name+当前日期作为excel名称保存。
    wbk.save(name)

# 如果该文件不是被import,则执行下面代码。
if __name__ == '__main__':
    # 定义一个字典，key为对应的数据类型也用作excel命名，value为查询语句
    sql = "select m.Name as pName,(select top 1 [Value] from WorkflowservicePlatformDB.dbo.v_ContentMemory where ProcessID = p.UID and [Key] = 'serialNumber0') as serialNumber,cm.[Value] ,ISNULL(dv.OrgName,org.OrgName) OrgName,p.InitiatorName,t.Executor, CASE p.Status when 'completed' then '结束' else STUFF((SELECT ',' + Label +':'+ ExecutorName FROM WorkflowservicePlatformDB.dbo.WorkQueue  w where ProcessID = p.UID and Priority = 0 and p.Status = 'running' AND w.Finished is null for xml path('')),1,1,'') end,p.Name,p.Created as [CreateDate],p.Status,p.Finished from WorkflowservicePlatformDB.dbo.Processes p join WorkflowservicePlatformDB.dbo.Models m on p.ModelID = m.UID left join WorkflowservicePlatformDB.dbo.v_ContentMemory cm on cm.ProcessID = p.UID and cm.[Key] = 'rdlType_Value' and m.Alias = 'EIAC-OA-BIZ-7CCFE5' left join WorkflowservicePlatformDB.dbo.v_Tracking t on t.ActivityID in (select a.[UID] from WorkflowservicePlatformDB.dbo.Activities a  join WorkflowservicePlatformDB.dbo.Models  m on m.UID = a.ModelID  where a.Name = 'Custom_02' and m.Alias = 'EIAC-OA-BIZ-AB0C6F') and t.ProcessID = p.UID LEFT JOIN  EIAC_UUV.dbo.v_Org_Dept_Relation dv ON dv.cOrgID = p.AcDeptID JOIN EIAC_UUV.dbo.UV_ORG org ON org.OrgID = p.AcDeptID where p.Status <> 'canceled' and  Alias in ('EIAC-OA-BIZ-7CCFE5','EIAC-OA-BIZ-AB0C6F') order by m.Name,CreateDate desc"
    db_dict = {py_mssql().file_name(): sql}
    # 遍历字典每个元素的key和value。
    for k, v in db_dict.items():
        # 用字典的每个key和value调用write_data_to_excel函数。
        write_data_to_excel(k, v)
