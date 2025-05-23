import openpyxl
from openpyxl.styles import Font
import os

# 获取当前脚本所在目录
base_dir = os.path.dirname(os.path.abspath(__file__))
input_path = os.path.join(base_dir, '1.xlsx')
output_path = os.path.join(base_dir, 'output_py.xlsx')

# 加载Excel文件
wb = openpyxl.load_workbook(input_path)
ws = wb.active  # 默认第一个sheet

# 设置表头（第一行）字体：偶数列为绿色，其余为红色
for idx, cell in enumerate(ws[1], start=1):
    if idx % 2 == 0:
        cell.font = Font(color='00B050')  # 绿色
    else:
        cell.font = Font(color='FF0000')  # 红色

# 保存新文件
wb.save(output_path)
print('处理完成，表头偶数列为绿色，其余为红色！（Python版）') 