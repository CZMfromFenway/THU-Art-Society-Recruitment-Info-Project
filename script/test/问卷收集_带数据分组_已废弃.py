import requests
import time
import pandas as pd
from datetime import datetime
import os
import hashlib

def manual_cookie_exporter():
    # 在这里粘贴你获取到的Cookie字符串
    cookie_str = input("请粘贴完整的Cookie字符串: ").strip()
    
    # 解析Cookie
    cookies = {}
    for item in cookie_str.split(';'):
        item = item.strip()
        if '=' in item:
            key, value = item.split('=', 1)
            cookies[key] = value
    
    # 设置请求参数
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Referer': 'https://www.wjx.cn/',
        'Accept': 'application/vnd.ms-excel'
    }
    
    url = "https://www.wjx.cn/wjx/activitystat/viewstatsummary.aspx?activity=331201067&reportid=0&dw=1&dt=0"
    output_file = "问卷数据.xlsx"
    
    session = requests.Session()
    session.cookies.update(cookies)
    
    # 创建数据指纹集合，用于去重
    existing_hashes = set()
    
    # 如果已有数据文件，加载现有数据的指纹
    if os.path.exists(output_file):
        try:
            existing_data = pd.read_excel(output_file)
            print(f"发现现有数据文件，共 {len(existing_data)} 条记录")
            
            # 为现有数据生成指纹
            for _, row in existing_data.iterrows():
                # 排除时间戳列生成指纹
                row_without_time = row.drop('导出时间') if '导出时间' in row else row
                row_hash = generate_row_hash(row_without_time)
                existing_hashes.add(row_hash)
                
            print(f"已加载 {len(existing_hashes)} 个数据指纹用于去重")
        except Exception as e:
            print(f"读取现有数据文件时出错: {e}")
    
    print("开始定时导出数据(每5秒)...")
    print("按 Ctrl+C 停止")
    
    count = 0
    new_data_count = 0
    duplicate_count = 0
    
    try:
        while True:
            count += 1
            print(f"\n第 {count} 次尝试导出...")
            
            response = session.get(url, headers=headers, timeout=10)
            
            # 检查是否是Excel文件
            content_type = response.headers.get('Content-Type', '').lower()
            if 'excel' in content_type or 'spreadsheet' in content_type:
                # 处理Excel文件
                temp_file = f"temp_{datetime.now().strftime('%H%M%S')}.xlsx"
                with open(temp_file, 'wb') as f:
                    f.write(response.content)
                
                # 读取新数据
                new_data = pd.read_excel(temp_file)
                new_data['导出时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                
                # 过滤重复数据
                unique_new_data = filter_duplicates(new_data, existing_hashes)
                
                if len(unique_new_data) > 0:
                    # 追加到主文件
                    if os.path.exists(output_file):
                        existing_data = pd.read_excel(output_file)
                        combined_data = pd.concat([existing_data, unique_new_data], ignore_index=True)
                        combined_data.to_excel(output_file, index=False)
                    else:
                        unique_new_data.to_excel(output_file, index=False)
                    
                    # 更新指纹集合
                    for _, row in unique_new_data.iterrows():
                        row_without_time = row.drop('导出时间')
                        row_hash = generate_row_hash(row_without_time)
                        existing_hashes.add(row_hash)
                    
                    new_data_count += len(unique_new_data)
                    print(f"✓ 新增 {len(unique_new_data)} 条记录 - {datetime.now().strftime('%H:%M:%S')}")
                    
                    # 分组处理新数据
                    group_data_by_preference(unique_new_data)
                else:
                    duplicate_count += len(new_data)
                    print(f"○ 无新数据，跳过 {len(new_data)} 条重复记录")
                
                os.remove(temp_file)
                
            else:
                print("× 响应不是Excel文件")
                print(f"状态码: {response.status_code}")
                print(f"内容类型: {response.headers.get('Content-Type')}")
            
            # 显示统计信息
            print(f"累计: 新增 {new_data_count} 条, 跳过 {duplicate_count} 条重复数据")
            
            # 等待5秒
            for i in range(5, 0, -1):
                print(f"\r下次尝试: {i}秒后", end='', flush=True)
                time.sleep(1)
            print('\r', end='', flush=True)
            
    except KeyboardInterrupt:
        print(f"\n\n程序已停止。总共处理: {count} 次导出")
        print(f"新增数据: {new_data_count} 条")
        print(f"跳过重复: {duplicate_count} 条")
        
        # 最终分组处理所有数据
        if os.path.exists(output_file):
            all_data = pd.read_excel(output_file)
            group_data_by_preference(all_data, is_final=True)
            
    except Exception as e:
        print(f"发生错误: {e}")

def generate_row_hash(row):
    """
    为数据行生成唯一哈希值，用于去重
    """
    # 将行数据转换为字符串并生成MD5哈希
    row_str = ''.join([str(x) for x in row.values if pd.notna(x)])
    return hashlib.md5(row_str.encode('utf-8')).hexdigest()

def filter_duplicates(new_data, existing_hashes):
    """
    过滤掉已经存在的数据行
    """
    unique_rows = []
    
    for _, row in new_data.iterrows():
        # 排除时间戳列进行去重比较
        row_without_time = row.drop('导出时间') if '导出时间' in row else row
        row_hash = generate_row_hash(row_without_time)
        
        if row_hash not in existing_hashes:
            unique_rows.append(row)
            existing_hashes.add(row_hash)  # 立即添加到集合中避免重复比较
    
    if unique_rows:
        return pd.DataFrame(unique_rows)
    else:
        return pd.DataFrame()

def group_data_by_preference(data, is_final=False):
    """
    根据报名者的志愿顺序将数据分成五个组
    """
    if data.empty:
        return
    
    # 定义组名和对应的列名（根据PDF中的结构）
    groups = {
        '书法组': 'N',  # 第N列应该是书法组志愿
        '国画组': 'O',  # 第O列应该是国画组志愿
        '西画组': 'P',  # 第P列应该是西画组志愿
        '漫画组': 'Q',  # 第Q列应该是漫画组志愿
        '篆刻组': 'R'   # 第R列应该是篆刻组志愿
    }
    
    # 创建输出目录
    output_dir = "分组数据"
    os.makedirs(output_dir, exist_ok=True)
    
    # 为每个组创建数据框
    group_dataframes = {group_name: pd.DataFrame() for group_name in groups.keys()}
    
    # 处理每一行数据
    for _, row in data.iterrows():
        # 找出第一志愿
        first_preference = None
        for group_name, col_name in groups.items():
            try:
                # 检查该列的值是否为1（第一志愿）
                if pd.notna(row[col_name]) and int(row[col_name]) == 1:
                    first_preference = group_name
                    break
            except (ValueError, KeyError):
                continue
        
        # 如果有第一志愿，添加到对应组
        if first_preference:
            group_dataframes[first_preference] = pd.concat([
                group_dataframes[first_preference], 
                row.to_frame().T
            ], ignore_index=True)
    
    # 保存每个组的数据到Excel文件
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    for group_name, df in group_dataframes.items():
        if not df.empty:
            filename = f"{output_dir}/{group_name}_报名表_{timestamp}.xlsx"
            df.to_excel(filename, index=False)
            print(f"✓ 已保存 {group_name} 数据: {len(df)} 条记录")
    
    # 如果是最终处理，也保存汇总文件
    if is_final:
        summary_filename = f"{output_dir}/所有分组汇总_{timestamp}.xlsx"
        with pd.ExcelWriter(summary_filename) as writer:
            for group_name, df in group_dataframes.items():
                if not df.empty:
                    df.to_excel(writer, sheet_name=group_name, index=False)
        print(f"✓ 已保存所有分组汇总文件")

# 运行程序
if __name__ == "__main__":
    manual_cookie_exporter()