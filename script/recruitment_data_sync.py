import requests
import time
import pandas as pd
from datetime import datetime
import os
import json
import argparse

import collect_raw_data
import parse_raw_data
import uploader

class RecruitmentDataSync:

    wjx_url: str
    wjx_cookie: str
    raw_data_file: str
    grouped_data_dir: str
    feishu_token: str
    uploader = uploader.Uploader()
    period = 500

    def __init__(self, wjx_url, wjx_cookie, raw_data_file, grouped_data_dir, feishu_token, preiod = 500):
        self.wjx_url = wjx_url
        self.wjx_cookie = wjx_cookie
        self.raw_data_file = raw_data_file
        self.grouped_data_dir = grouped_data_dir
        self.feishu_token = feishu_token
        self.period = preiod

    def start(self, reset = False):

        if reset:
            self.feishu_token = self.uploader.reset_all_sheets(self.feishu_token)
            print("重置飞书表格完成")
            if os.path.exists(self.raw_data_file):
                os.remove(self.raw_data_file)
                print("删除本地原始数据完成")
            if os.path.exists(self.grouped_data_dir):
                for file in os.listdir(self.grouped_data_dir):
                    os.remove(os.path.join(self.grouped_data_dir, file))
                print("删除本地分组数据完成")

        output_file = self.raw_data_file
        
        # 解析Cookie
        cookies = {}
        for item in self.wjx_cookie.split(';'):
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
                    row_hash = collect_raw_data.generate_row_hash(row_without_time)
                    existing_hashes.add(row_hash)
                    
                print(f"已加载 {len(existing_hashes)} 个数据指纹用于去重")
            except Exception as e:
                print(f"读取现有数据文件时出错: {e}")
        
        print("开始定时导出数据(每5秒)...")
        print("按 Ctrl+C 停止")

        count = 0
        data_count = 0
        duplicate_count = 0
        last_time = datetime.min

        try:
            while True:
                count += 1
                print(f"\n第 {count} 次尝试导出...")
                
                new_data_count, duplicate_count = collect_raw_data.export(session, self.wjx_url, headers, output_file, existing_hashes, data_count, duplicate_count)
                
                if new_data_count > data_count:
                    data_count = new_data_count
                    # 数据归类
                    time_str = parse_raw_data.process_recruitment_data(output_file, self.grouped_data_dir, last_time)
                    last_time = datetime.now()
                    # 上传飞书
                    self.feishu_token = self.uploader.parse_excel(self.feishu_token, self.grouped_data_dir, time_str)
                
                # 等待
                for i in range(self.period, 0, -1):
                    print(f"\r下次尝试: {i}秒后", end='', flush=True)
                    time.sleep(1)
                print('\r', end='', flush=True)
                
        except KeyboardInterrupt:
            print(f"\n\n程序已停止。总共处理: {count} 次导出")
            print(f"新增数据: {new_data_count} 条")
            print(f"跳过重复: {duplicate_count} 条")
        except Exception as e:
            print(f"发生错误: {e}")

if __name__ == "__main__":

    # 根据json配置参数
    with open('config/config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
        wjx_url = config.get('wjx_url')  # 替换为实际问卷链接
        wjx_cookie = config.get('wjx_cookie')  # 替换为实际Cookie
        raw_data_file = config.get('raw_data_file', 'raw_data\问卷数据.xlsx')
        grouped_data_dir = config.get('grouped_data_dir', 'grouped_data')
        feishu_token = config.get('feishu_token')  # 替换为实际飞书API Token
        period = config.get('period', 500)
    
    # 命令行参数：--reset / -r 指定是否重置之前的数据（默认 False）
    parser = argparse.ArgumentParser(description="Recruitment data sync runner")
    parser.add_argument('-r', '--reset', action='store_true', help='重置飞书表格和本地数据（默认否）')
    args = parser.parse_args()

    recruitment = RecruitmentDataSync(wjx_url, wjx_cookie, raw_data_file, grouped_data_dir, feishu_token, period)
    recruitment.start(args.reset)