import pandas as pd
import os

def convert_xlsx_to_csv():
    try:
        # 获取当前目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(current_dir)
        
        # 输入和输出路径
        input_path = os.path.join(project_root, 'data', 'input', '2023年选课结果.xlsx')
        output_path = os.path.join(project_root, 'data', 'input', '2023年选课结果.csv')
        
        # 检查文件是否存在
        if not os.path.exists(input_path):
            print(f"错误：找不到输入文件 {input_path}")
            return
            
        # 读取xlsx
        print("正在读取Excel文件...")
        df = pd.read_excel(input_path)
        
        # 保存为csv
        print("正在转换为CSV文件...")
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
        
        # 输出调剂录取的学生名单
        print("\n=== 调剂录取学生名单 ===")
        print("以下学生的第一志愿选择与最终录取结果不同：\n")
        
        # 专业选项对应表
        major_choices = {
            'a': ['电子信息工程', '通信工程', '电磁场与无线技术'],
            'b': ['电子信息工程', '电磁场与无线技术', '通信工程'],
            'c': ['电磁场与无线技术', '电子信息工程', '通信工程'],
            'd': ['电磁场与无线技术', '通信工程', '电子信息工程'],
            'e': ['通信工程', '电子信息工程', '电磁场与无线技术'],
            'f': ['通信工程', '电磁场与无线技术', '电子信息工程']
        }
        
        # 计数器
        adjustment_count = 0
        
        # 遍历每个学生
        for _, row in df.iterrows():
            if pd.isna(row['选课选项']):
                continue
                
            choice = row['选课选项'].lower()
            if choice in major_choices:
                first_choice = major_choices[choice][0]
                final_result = row['最终结果']
                
                # 标准化最终结果
                if '电信' in final_result:
                    final_result = '电子信息工程'
                elif '通信' in final_result:
                    final_result = '通信工程'
                elif '电磁' in final_result:
                    final_result = '电磁场与无线技术'
                    
                if first_choice != final_result:
                    adjustment_count += 1
                    print(f"姓名：{row['姓名']}")
                    print(f"班级：{row['班级']}")
                    print(f"成绩：{row['成绩']}")
                    print(f"第一志愿：{first_choice}")
                    print(f"最终录取：{final_result}")
                    print("-" * 30)
        
        print(f"\n总共有 {adjustment_count} 名学生被调剂录取")
        
    except Exception as e:
        print(f"发生错误：{str(e)}")

if __name__ == '__main__':
    convert_xlsx_to_csv() 