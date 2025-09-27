import requests
import time
import pandas as pd
from datetime import datetime
import os
import hashlib

# url是导出excel表格的链接，第30行
# cookie是问卷星登录后主页面的，具体获得流程为F12————network————刷新后第一条请求————请求标头————COOKIE

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
    
    url = "https://www.wjx.cn/wjx/activitystat/viewstatsummary.aspx?activity=330193666&reportid=0&dw=1&dt=2"
    output_file = "raw_data\问卷数据.xlsx"
    
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

# 更精确的去重版本（如果需要更严格的去重）
def advanced_filter_duplicates(new_data, existing_data=None):
    """
    更精确的去重方法，比较所有列的值
    """
    if existing_data is None or existing_data.empty:
        return new_data
    
    # 找出新数据中不在现有数据中的行
    # 首先确保列名一致
    common_columns = [col for col in new_data.columns if col in existing_data.columns and col != '导出时间']
    
    if not common_columns:
        return new_data
    
    # 合并数据并标记重复
    merged = pd.concat([existing_data[common_columns], new_data[common_columns]])
    duplicates = merged.duplicated(keep='first')
    
    # 只保留新数据中不重复的行
    new_data_indices = range(len(existing_data), len(merged))
    unique_mask = [not duplicates[i] for i in new_data_indices]
    
    return new_data[unique_mask]

# 运行程序
if __name__ == "__main__":
    manual_cookie_exporter()