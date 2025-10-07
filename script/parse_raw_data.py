import pandas as pd
import os
from datetime import datetime
import re

def find_interview_columns(df, group_name):
    """
    自动识别指定组别的面试时间列
    返回包含6个时间段的列名列表
    """
    # 查找包含组名和"面试的时间"的列
    pattern = f"您能参加【{group_name}】面试的时间"
    interview_columns = []
    
    for col in df.columns:
        if pattern in str(col):
            # 找到起始列，获取其后5列（共6个时间段）
            col_index = df.columns.get_loc(col)
            interview_columns = list(df.columns[col_index:col_index+6])
            break
    
    return interview_columns

def map_column_to_time_label(column_name):
    """
    将原始列名映射为简化的时间标签
    """
    if '周三' in column_name and '14：00' in column_name:
        return '周三下午'
    elif '周三' in column_name and '18：00' in column_name:
        return '周三晚上'
    elif '周四' in column_name and '14：00' in column_name:
        return '周四下午'
    elif '周四' in column_name and '18：00' in column_name:
        return '周四晚上'
    elif '周五' in column_name and '14：00' in column_name:
        return '周五下午'
    elif '周五' in column_name and '18：00' in column_name:
        return '周五晚上'
    else:
        # 如果无法识别，返回原始列名的简化版本
        return re.sub(r'[^周一二三四五六日下午晚上]', '', column_name)

def create_highlighted_excel(df, output_filename):
    """
    创建带有条件格式的Excel文件，高亮显示面试时间选择
    """
    # 使用xlsxwriter引擎来支持条件格式
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='面试信息', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['面试信息']
        
        # 定义高亮格式
        highlight_format = workbook.add_format({
            'bg_color': '#FFD700',  # 黄色背景
            'bold': True
        })
        
        # 确定面试时间列的起始位置（假设面试时间列在报名志愿之后）
        time_columns = ['周三下午', '周三晚上', '周四下午', '周四晚上', '周五下午', '周五晚上']
        
        # 应用条件格式到面试时间列
        for i, col_name in enumerate(time_columns):
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name)
                # 跳过表头，从第二行开始
                worksheet.conditional_format(1, col_idx, len(df), col_idx, {
                    'type': 'cell',
                    'criteria': '==',
                    'value': 1,
                    'format': highlight_format
                })

def process_recruitment_data(file_path, cutoff_time = datetime.min):
    """
    处理美术社招新报名表，按组别筛选并导出面试信息
    
    :param file_path: 报名表文件路径
    :param cutoff_time: 筛选提交时间晚于此时间的记录，默认为很久以前
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        print(f"成功读取文件，共{len(df)}条记录")
        print("列名预览:")
        for i, col in enumerate(df.columns):
            print(f"{i}: {col}")
        
        # 将“提交答卷时间”列转换为datetime类型
        if '提交答卷时间' in df.columns:
            df['提交答卷时间'] = pd.to_datetime(df['提交答卷时间'], errors='coerce')
        else:
            print("错误：未找到'提交答卷时间'列")
            return
        
        # 筛选提交时间晚于cutoff_time的记录
        df = df[df['提交答卷时间'] > cutoff_time]
        print(f"筛选后剩余记录数: {len(df)}")
        
        # 定义组别映射
        groups = {
            '书法组': '2、(书法组)',
            '国画组': '2、(国画组)', 
            '篆刻组': '2、(篆刻组)',
            '西画组': '2、(西画组)',
            '漫画组': '2、(漫画组)'
        }
        
        # 宣传技能列
        promotion_skill_columns = [
            '13、是否有宣传相关技能(推送排版)',
            '13、(平面设计)',
            '13、(视频制作)',
            '13、(文案写作)',
            '13、(其他)'
        ]
        
        promotion_interest_column = '14、是否有兴趣加入美社宣传小组，接受宣传技能培训，参与平面设计、专栏采写、文创IP策划等业务'
        
        # 处理每个组别
        for group_name, priority_column in groups.items():
            print(f"\n正在处理{group_name}...")
            
            # 自动识别面试时间列
            interview_columns = find_interview_columns(df, group_name)
            
            if not interview_columns:
                print(f"  警告：未找到{group_name}的面试时间列，跳过该组")
                continue
                
            if len(interview_columns) < 6:
                print(f"  警告：{group_name}的面试时间列不足6列，找到{len(interview_columns)}列")
            
            # 创建时间列映射
            time_mapping = {}
            for orig_col in interview_columns[:6]:  # 只取前6列
                time_label = map_column_to_time_label(orig_col)
                time_mapping[orig_col] = time_label
            
            print(f"  {group_name}面试时间列映射: {time_mapping}")
            
            # 筛选该组有志愿的报名者（志愿列为1-5的数字）
            group_df = pd.DataFrame()  # Initialize an empty DataFrame
            for _, row in df.iterrows():
                if pd.notna(row[priority_column]):
                    group_df = pd.concat([group_df, row.to_frame().T], ignore_index=True)
            
            if len(group_df) == 0:
                print(f"  {group_name}没有报名者")
                continue
                
            print(f"  {group_name}共有{len(group_df)}名报名者")
            
            # 创建新的数据框
            result_data = []
            
            for _, row in group_df.iterrows():
                # 提取基本信息
                basic_info = {
                    '提交时间': row['提交答卷时间'],
                    '姓名': row['1、您的基本信息—姓名：'],
                    '院系': row['1、院系：'],
                    '班级': row['1、班级:'],
                    '手机': row['1、手机号：'],
                    '微信号': row['1、微信号：'],
                    '报名志愿': row[priority_column]
                }
                
                # 添加面试时间列（6列，0或1）
                for orig_col, time_label in time_mapping.items():
                    if orig_col in row and pd.notna(row[orig_col]):
                        basic_info[time_label] = int(row[orig_col])
                    else:
                        basic_info[time_label] = 0
                
                # 专业水平概况（留空）
                basic_info['专业水平概况'] = ''
                
                # 处理宣传意愿和技能
                promotion_skills = []
                for skill_col in promotion_skill_columns:
                    if skill_col in row and pd.notna(row[skill_col]) and row[skill_col] == 1:
                        skill_name = skill_col.split('(')[-1].replace(')', '') if '(' in skill_col else skill_col
                        promotion_skills.append(skill_name)
                
                promotion_info = []
                if promotion_skills:
                    promotion_info.append('技能：' + '、'.join(promotion_skills))
                
                if promotion_interest_column in row and pd.notna(row[promotion_interest_column]):
                    interest_value = row[promotion_interest_column]
                    if interest_value == 1:
                        promotion_info.append('有兴趣加入宣传小组')
                    elif interest_value == 2:
                        promotion_info.append('无兴趣加入宣传小组')
                
                basic_info['宣传意愿和技能'] = '；'.join(promotion_info) if promotion_info else '无'
                
                result_data.append(basic_info)
            
            # 创建结果DataFrame
            result_df = pd.DataFrame(result_data)
            
            # 确保时间列的固定顺序
            time_order = ['周三下午', '周三晚上', '周四下午', '周四晚上', '周五下午', '周五晚上']
            existing_time_cols = [col for col in time_order if col in result_df.columns]
            
            # 重新排列列的顺序
            base_columns = ['提交时间', '姓名', '院系', '班级', '手机', '微信号', '报名志愿']
            other_columns = [col for col in result_df.columns if col not in base_columns + existing_time_cols]
            
            final_columns = base_columns + existing_time_cols + other_columns
            result_df = result_df[final_columns]
            
            # 保存为带有条件格式的Excel文件
            output_filename = f"grouped_data\{group_name}面试信息_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            create_highlighted_excel(result_df, output_filename)
            print(f"  {group_name}面试信息已保存为: {output_filename}")
            
            # 显示前几行作为预览
            print(f"  数据预览（前3行）:")
            print(result_df.head(3).to_string(index=False))
            
    except Exception as e:
        print(f"处理过程中出现错误: {e}")
        import traceback
        traceback.print_exc()

def main():
    """
    主函数
    """
    # 文件路径 - 请根据实际情况修改
    file_path = "问卷数据.xlsx"
    
    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"文件 {file_path} 不存在，请检查文件路径")
        return
    
    # 处理数据
    process_recruitment_data(file_path)
    print("\n处理完成！")

if __name__ == "__main__":
    main()