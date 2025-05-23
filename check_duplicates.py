import pandas as pd
import sys
import argparse
import os

def check_column_duplicates(file_path, column=None):
    """
    检查Excel文件指定列的重复项
    
    参数:
    file_path: Excel文件路径
    column: 列名或列索引（例如'A'或0代表第一列，'B'或1代表第二列，以此类推）
            如果不指定，则默认检查第一列(A列)
    
    返回:
    字符串，表示检查结果
    """
    # 读取Excel文件
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        return f"读取Excel文件出错: {str(e)}"
    
    # 检查文件是否为空
    if len(df.columns) == 0:
        return "Excel文件为空或没有列"
    
    # 确定要检查的列
    col_data = None
    col_name = "A列"  # 默认显示名称
    
    # 根据用户输入确定列
    if column is None:
        # 默认使用第一列(A列)
        col_data = df.iloc[:, 0]
    elif isinstance(column, int):
        # 使用列索引（数字）
        if 0 <= column < len(df.columns):
            col_data = df.iloc[:, column]
            col_name = f"列索引{column}({chr(65 + column)}列)"
        else:
            return f"列索引{column}超出范围，文件仅包含{len(df.columns)}列"
    elif isinstance(column, str):
        # 处理列字母（如'A','B'等）
        if len(column) == 1 and 'A' <= column.upper() <= 'Z':
            col_idx = ord(column.upper()) - ord('A')
            if col_idx < len(df.columns):
                col_data = df.iloc[:, col_idx]
                col_name = f"{column.upper()}列"
            else:
                return f"{column.upper()}列超出范围，文件仅包含{len(df.columns)}列"
        # 尝试将输入作为列名处理
        elif column in df.columns:
            col_data = df[column]
            col_name = f"列名'{column}'"
        else:
            return f"找不到列'{column}'，请检查列名或使用列字母(A-Z)"
    
    # 检查重复值
    duplicates = col_data[col_data.duplicated()]
    
    if duplicates.empty:
        return f"{col_name}中没有重复项"
    else:
        # 获取所有重复的值及其所有出现的位置（行号）
        result = []
        duplicate_values = set(duplicates)
        for value in duplicate_values:
            # 找出该值在列中出现的所有行位置（行号从1开始，符合Excel习惯）
            rows = [i + 1 for i, v in enumerate(col_data) if v == value]
            if len(rows) > 1:  # 确保至少有2行才算重复
                result.append(f"值 '{value}' 在{col_name}中重复出现，行号为: {rows}")
        
        return "\n".join(result)

def interactive_mode():
    """交互式模式，引导用户完成Excel重复项检查"""
    print("="*50)
    print("Excel重复项检查工具")
    print("="*50)
    
    # 获取当前目录下所有Excel文件
    excel_files = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls'))]
    
    if not excel_files:
        print("当前目录下没有找到Excel文件(.xlsx或.xls)")
        file_path = input("请输入Excel文件的完整路径: ")
    else:
        # 显示可选的Excel文件
        print("\n当前目录下的Excel文件:")
        for i, file in enumerate(excel_files, 1):
            print(f"{i}. {file}")
        
        # 让用户选择文件
        choice = input("\n请选择要检查的Excel文件序号（或输入其他文件路径）: ")
        try:
            index = int(choice) - 1
            if 0 <= index < len(excel_files):
                file_path = excel_files[index]
            else:
                print("无效的选择，请输入有效的序号")
                return
        except ValueError:
            # 用户可能直接输入了文件路径
            file_path = choice
    
    # 确保文件存在
    if not os.path.exists(file_path):
        print(f"错误：文件 '{file_path}' 不存在")
        return
    
    try:
        # 读取文件以获取列信息
        df = pd.read_excel(file_path)
        num_columns = len(df.columns)
        
        print(f"\n文件 '{file_path}' 包含 {num_columns} 列")
        print("列选项:")
        
        # 显示列名和对应的字母
        for i, col_name in enumerate(df.columns):
            col_letter = chr(65 + i) if i < 26 else f"A{chr(65 + i - 26)}"  # 处理超过Z的列
            print(f"{col_letter}. {col_name}")
        
        # 让用户选择列
        column_choice = input("\n请输入要检查的列字母（如A、B、C...）或直接按回车检查A列: ")
        if not column_choice:
            column_choice = "A"  # 默认检查A列
        
        # 执行检查
        result = check_column_duplicates(file_path, column_choice)
        
        print("\n===== 检查结果 =====")
        print(result)
        
        # 询问是否保存结果到文件
        save_choice = input("\n是否将结果保存到文件？(y/n): ").lower()
        if save_choice == 'y':
            output_file = input("请输入保存文件名(默认 'duplicate_results.txt'): ")
            output_file = output_file if output_file else "duplicate_results.txt"
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(f"文件: {file_path}\n")
                f.write(f"检查列: {column_choice}\n\n")
                f.write(result)
            print(f"结果已保存到 {output_file}")
            
        # 询问是否继续检查其他列
        continue_choice = input("\n是否检查其他列？(y/n): ").lower()
        if continue_choice == 'y':
            # 递归调用，让用户选择新的列
            column_choice = input("\n请输入要检查的列字母（如A、B、C...）: ")
            if column_choice:
                result = check_column_duplicates(file_path, column_choice)
                print("\n===== 检查结果 =====")
                print(result)
    
    except Exception as e:
        print(f"处理文件时出错: {str(e)}")
    
    print("\n感谢使用Excel重复项检查工具！")

def main():
    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(description='检查Excel文件中的重复项')
    parser.add_argument('file', type=str, nargs='?', help='Excel文件路径')
    parser.add_argument('-c', '--column', type=str, help='要检查的列(例如: A, B, C...或列名)')
    parser.add_argument('-i', '--interactive', action='store_true', help='使用交互式模式')
    
    # 解析命令行参数
    args = parser.parse_args()
    
    # 如果指定了交互式模式或没有提供任何参数，则进入交互模式
    if args.interactive or (not args.file and len(sys.argv) == 1):
        interactive_mode()
    else:
        # 使用命令行模式
        file_path = args.file if args.file else "test.xlsx"  # 默认为test.xlsx
        
        # 如果提供了列参数，解析它
        column = None
        if args.column:
            # 尝试将输入转换为数字（列索引）
            try:
                column = int(args.column)
            except ValueError:
                # 不是数字，则按字母或列名处理
                column = args.column
        
        # 调用主函数
        result = check_column_duplicates(file_path, column)
        print(result)

if __name__ == "__main__":
    main() 