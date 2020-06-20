# 批量大小
# 安全库存
# 提前期
# 数据文件（period.dat）
# 各时段订单量（order.dat）
# 各时段预测量（prediction.dat）
# 各时段计划接收量（ScheduleReceipts.adt）
# 过去时段的预计可用库存（PrevInventory.dat）

# 1.计算所有时段毛需求量
# 2.按照从时段1到时段n的顺序计算n个时段毛需求量
# 3.以此计算该阶段净需求、计划产出量及预计可用库存
# 4.依次计算所有时段计划投入量
# 5.依次计算所有时段可供销售量

# by Jiefeng_Lin

import xlsxwriter


# 读取文件函数
def ReadMatInfo(filepath):
    with open(filepath, "r") as f:
        contents = f.read()
        tmp = contents.split(' ', -1)
        return tmp


# 从文件现有初始预计可用库存量
now_stock = []
now_stock.insert(0, int(ReadMatInfo('data_files/PrevInventory.dat')[1]))

# 从文件读取安全库存量
safe_stock = int(ReadMatInfo('data_files/matinfo.dat')[2])

# 从文件批量大小
product_batch = int(ReadMatInfo('data_files/matinfo.dat')[1])

# 从文件读取提前期
pre_date = int(ReadMatInfo('data_files/matinfo.dat')[3])

# 从文件读取总时段
total_period = int(ReadMatInfo('data_files/period.dat')[1])

# 从文件读取需求时界
require_period = int(ReadMatInfo('data_files/period.dat')[2])

# 从文件读取计划时界
plan_period = int(ReadMatInfo('data_files/period.dat')[3])

# 从文件读取预测量
pre = []
for i in range(total_period):
    pre.insert(i, int(ReadMatInfo('data_files/prediction.dat')[i + 1]))

# 从文件读取订单量
order = []
for i in range(total_period):
    order.insert(i, int(ReadMatInfo('data_files/order.dat')[i + 1]))

# 从文件读取计划接收量
schedule_receipt = []
for i in range(total_period):
    schedule_receipt.insert(i, int(ReadMatInfo('data_files/ScheduledReceipts.dat')[i + 1]))

# 定义净需求量数组
neet_reqiire = []
neet_reqiire.insert(0, 0)

# 定义毛需求量数组
gross_require = []

# 定义可供销售量数组
ATP = []
ATP.insert(0, 0)

# 定义计划产出量数组
plan_production = []
plan_production.insert(0, 0)

# 定义计划投入量数组
plan_release = []
plan_release.insert(total_period - 1, 0)

# 计算毛需求量
for i in range(total_period):
    # 当在需求时段，毛需求=订单量
    if i <= require_period:
        gross_require.insert(i, order[i])
    # 当在计划阶段，毛需求=max(订单量，预测量)
    if i > require_period and i <= plan_period:
        gross_require.insert(i, max(order[i], pre[i]))
    # 在预测阶段，毛需求=预测量
    if i > plan_period and i < total_period:
        gross_require.insert(i, pre[i])


# 定义批量大小增量的函数
def calculate_increse_of_product_batch(n):
    # n是大于或等于1的整数
    if n <= 0:
        return 0
    else:
        for i in range(1, 11):
            # 当 i*生产批量 >= 净需求量 > (i-1)*生产批量的时候，满足计划产出量计算要求
            if (i - 1) * product_batch < n and i * product_batch >= n:
                return i * product_batch
                # 计划产出量 = i * 生产批量


# 计算净需求量
for i in range(1, total_period):
    # 净需求 = 本时段毛需求量 + 安全库存 - 前时段可用库存 - 计划接收量
    neet_reqiire.insert(i, gross_require[i] + safe_stock - now_stock[i - 1] - schedule_receipt[i])
    # 计划产出直接应用上面的函数
    plan_production.insert(i, calculate_increse_of_product_batch(neet_reqiire[i]))
    # 预计可用库存 = 前时段预计可用库存 + 预计接收量 + 计划产出量 - 毛需求
    now_stock.insert(i, now_stock[i - 1] + schedule_receipt[i] + plan_production[i] - gross_require[i])

# 计算计划投入量
for i in range(total_period - 1):
    # 本时段计划投入量 = 下时段计划产出量
    plan_release.insert(i, plan_production[i + 1])

# 计算可供销售量ATP
for i in range(total_period - 1):
    # ATP = 本时段计划产出量 + 本时段计划接收量 + 上时段预计可用库存 - 本时段订单量
    if i == 1:
        ATP.insert(i, plan_production[i] + schedule_receipt[i] + now_stock[i - 1] - order[i])
    # ATP = 本时段计划产出量 + 本时段计划接收量 - 本时段订单量
    else:
        ATP.insert(i, plan_production[i] + schedule_receipt[i] - order[i])

# 将数据写入到Excel并创建表格
# 创建excel文档
print("正在进行excel录入中")
workbook = xlsxwriter.Workbook("./data_files/calculation_of_MPS1.xlsx")
# 创建一个sheet
worksheet = workbook.add_worksheet(name="calculation_of_MPS")
# 确定横轴纵轴属性
calculation_item = ['预测量', '订单量', '毛需求量', '计划接收量', '预计可用库存', '净需求量', '计划产出量', '计划投入量', '可供销售量']
headline = ['时区/计算类别', '过去时段', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10']
# 写入到表格中
worksheet.write_row("A1", headline)
worksheet.write_column("A2", calculation_item)
worksheet.write_row("B2", pre)
worksheet.write_row("B3", order)
worksheet.write_row("B4", gross_require)
worksheet.write_row("B5", schedule_receipt)
worksheet.write_row("B6", now_stock)
worksheet.write_row("B7", neet_reqiire)
worksheet.write_row("B8", plan_production)
worksheet.write_row("B9", plan_release)
worksheet.write_row("B10", ATP)
workbook.close()
print("excel录入成功")
# 在控制台输出数据：
print("预测量是：")
print(pre)
print("订单量是：")
print(order)
print("毛需求是：")
print(gross_require)
print("计划接收量是：")
print(schedule_receipt)
print("预计可用库存是：")
print(now_stock)
print("净需求量是：")
print(neet_reqiire)
print("计划产出是：")
print(plan_production)
print("计划投入是:")
print(plan_release)
print("可供销售是：")
print(ATP)