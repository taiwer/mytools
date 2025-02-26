import os
import pandas as pd
from gooey import Gooey, GooeyParser
import glob

@Gooey(program_name="MyTools",  # 修改窗口名称为 MyTools
       default_size=(600, 500),
       language='chinese',
       sidebar_title="功能",  # 添加侧边栏标题
       show_sidebar=True,
       disable_maximize=True, # 禁用窗口最大化按钮
       )     # 启用侧边栏
def main():
    # 创建解析器
    parser = GooeyParser(description="mytools")
    subs = parser.add_subparsers(help='step', dest='command')
    parser_1 = subs.add_parser('MergeExcel', help='MergeExcel')
    parser_2 = subs.add_parser('待开发', help='MergeExcel')
    
    # 添加参数
    parser_1.add_argument(
        'folder_path',
        metavar='文件夹路径',
        help='选择包含xlsx文件的文件夹',
        widget='DirChooser'  # 文件夹选择控件
    )
    
    parser_1.add_argument(
        '--start_column',
        metavar='起始列',
        help='数据开始的列数（从1开始计数）',
        default=2,
        widget='IntegerField',  # 整数输入框
        gooey_options={'min': 1, 'max': 100}
    )
    
    parser_1.add_argument(
        '--start_row',
        metavar='起始行',
        help='数据开始的行数（从1开始计数）',
        default=3,
        widget='IntegerField',  # 整数输入框
        gooey_options={'min': 1, 'max': 1000}
    )
    
    # 解析参数
    args = parser.parse_args()
    
    # 获取参数值
    folder_path = args.folder_path
    start_column = int(args.start_column) - 1  # 转换为0-based索引
    start_row = int(args.start_row) - 1  # 转换为0-based索引
    
    # 获取最低层文件夹名称
    folder_name = os.path.basename(folder_path)
    # 设置输出文件名
    output_file = os.path.join(os.path.dirname(folder_path), f"{folder_name}.xlsx")
    
    # 获取所有xlsx文件
    excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    
    if not excel_files:
        raise ValueError("The selected folder does not contain any xlsx files!")
    
    # 存储所有数据的列表
    all_data = []
    
    # 遍历每个xlsx文件
    for file_index, file_path in enumerate(excel_files):
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            
            # 检查列数是否足够
            if start_column >= len(df.columns):
                print(f"Error: No. {file_index} file's column is not enough, skip this file")
                continue
            
            # 获取文件名（不含路径和扩展名）
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            
            # 从指定行和列开始读取数据
            selected_data = df.iloc[start_row:, start_column]
            
            # 创建包含文件名和数据的DataFrame
            temp_df = pd.DataFrame({
                '文件名': [file_name] * len(selected_data),
                f'列{start_column + 1}': selected_data.values
            })
            
            # 添加到总数据列表
            all_data.append(temp_df)
            
        except Exception as e:
            print(f"Deal with No. {file_index} file error: {str(e)}")
    
    # 合并所有数据
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        
        # 保存到新的Excel文件
        final_df.to_excel(output_file, index=False)
        print("finished !!!")
    else:
        print("No data merged, please check the input parameters and file content")

if __name__ == '__main__':
    main()