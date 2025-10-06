import pandas as pd
import http.client
import json
import os

class Uploader:

    spreadsheet_token = {"书法组": ["IXrnsAVd6h9u2Lt9Tt7cOwEonzg", "01171e"],
                        "国画组": ["EKIosEUJLhc2Z8tvzoecJjCEnWe", "01171e"],
                        "篆刻组": ["EaE0ww3yQiNncbkiqnkcXgjVncd", "81776e"],
                        "西画组": ["Jp1TsCc5QhwaLttc5jvcZwQAnxc", "01171e"],
                        "漫画组": ["Sm0KsWsQohwxrztfAtAcJmDRnSc", "01171e"]}

    app_id = "cli_a86924444cb8d00c"
    app_secret = "ydkzWor6TkJGTa7WZrtHBUfGJ1mNWmn0"

    def __upload_to_feishu(authorization_token, data, group_name):
        if group_name not in Uploader.spreadsheet_token:
            print(f"未知小组: {group_name}")
            return

        spreadSheetToken, sheetId = Uploader.spreadsheet_token[group_name]
        path = f"/open-apis/sheets/v2/spreadsheets/{spreadSheetToken}/values_append?insertDataOption=INSERT_ROWS"

        # 准备数据
        values = data.values.tolist()

        # 将其中的nan替换为空字符串

        for r in range(len(values)):
            for c in range(len(values[r])):
                if pd.isna(values[r][c]):
                    values[r][c] = ""
        
        conn = http.client.HTTPSConnection("open.feishu.cn")

        payload = {
            "valueRange": {
                "range": f"{sheetId}!A2:{chr(64 + len(data.columns))}{len(data) + 1}",
                "values": values
            }
        }

        headers = {
            'Authorization': authorization_token,
            'Content-Type': 'application/json'
        }

        conn.request("POST", path, json.dumps(payload), headers)
        response = conn.getresponse()
        # 读取并解析响应（response.read() 返回 bytes，需 decode->json）
        resp_bytes = response.read()
        try:
            resp_text = resp_bytes.decode('utf-8')
        except Exception:
            resp_text = resp_bytes.decode(errors='replace')

        try:
            resp_json = json.loads(resp_text)
        except Exception:
            print(f"上传返回非JSON响应: HTTP {response.status} {response.reason}\n{resp_text}")
            return

        # 处理飞书返回结构：{"code":0, "msg":"success", "data": {...}}
        if resp_json.get('code') == 0:
            print(f"成功上传数据到 {group_name} 小组，响应: {resp_json}")
        else:
            print(f"上传失败: HTTP {response.status} {response.reason}, 响应: {resp_json}")

    def __delete_data(authorization_token, group_name, row):
        if group_name not in Uploader.spreadsheet_token:
            print(f"未知小组: {group_name}")
            return

        spreadSheetToken, sheedId = Uploader.spreadsheet_token[group_name]
        path = f"/open-apis/sheets/v2/spreadsheets/{spreadSheetToken}/dimension_range"

        conn = http.client.HTTPSConnection("open.feishu.cn")

        headers = {
            'Authorization': authorization_token,
            'Content-Type': 'application/json'
        }

        payload = json.dumps({
            "dimension": {
                "sheetId": sheedId,
                "majorDimension": "ROWS",
                "startIndex": 2,
                "endIndex": row
            }
        })

        conn.request("DELETE", path, payload, headers)
        response = conn.getresponse()
        resp_bytes = response.read()
        try:
            resp_text = resp_bytes.decode('utf-8')
        except Exception:
            resp_text = resp_bytes.decode(errors='replace')

        try:
            resp_json = json.loads(resp_text)
        except Exception:
            print(f"删除返回非JSON响应: HTTP {response.status} {response.reason}\n{resp_text}")
            return

        if resp_json.get('code') == 0:
            print(f"成功删除 {group_name} 小组的所有数据，响应: {resp_json}")
        else:
            print(f"删除失败: HTTP {response.status} {response.reason}, 响应: {resp_json}")

    def __get_sheet_rows(authorization_token, group_name):
        if group_name not in Uploader.spreadsheet_token:
            print(f"未知小组: {group_name}")
            return 0

        spreadSheetToken, sheedId = Uploader.spreadsheet_token[group_name]

        path = f"/open-apis/sheets/v3/spreadsheets/{spreadSheetToken}/sheets/{sheedId}"

        conn = http.client.HTTPSConnection("open.feishu.cn")

        payload = ''

        headers = {
            'Authorization': authorization_token
        }

        conn.request("GET", path, payload, headers)
        response = conn.getresponse()
        resp_bytes = response.read()
        try:
            resp_text = resp_bytes.decode('utf-8')
        except Exception:
            resp_text = resp_bytes.decode(errors='replace')

        try:
            resp_json = json.loads(resp_text)
        except Exception:
            print(f"获取行数返回非JSON响应: HTTP {response.status} {response.reason}\n{resp_text}")
            return 0

        if resp_json.get('code') == 0:
            # 优先从 merges 列表中取第一个合并区块的 end_row_index（如果存在）
            merges = resp_json.get('data', {}).get('sheet', {}).get('merges', [])
            if isinstance(merges, list) and len(merges) > 0 and isinstance(merges[0], dict):
                end_row_index = merges[0].get('end_row_index')
                if isinstance(end_row_index, int):
                    return end_row_index  # 返回 end_row_index（根据你的需求可减去表头）
            # 否则回退到 grid_properties 的 row_count
            row_count = resp_json.get('data', {}).get('sheet', {}).get('grid_properties', {}).get('row_count')
            if isinstance(row_count, int):
                return row_count
            # 无法获取时返回0
            return 0
        else:
            print(f"获取行数失败: HTTP {response.status} {response.reason}, 响应: {resp_json}")
            return 0

    def parse_excel(authorization_token, dir, time_str):
        for catagory in Uploader.spreadsheet_token.keys():
            file_path = f"{dir}/{catagory}面试信息_{time_str}.xlsx"
            if os.path.exists(file_path):
                data = pd.read_excel(file_path)
                Uploader.upload_to_feishu(authorization_token, data, catagory)

        return Uploader.get_tanent_access_token(authorization_token)
            
    def reset_all_sheets(authorization_token):
        for catagory in Uploader.spreadsheet_token.keys():
            rows = Uploader.get_sheet_rows(catagory)
            print(f"正在清空 {catagory} 小组数据，共 {rows - 1} 行")
            if rows > 1:
                Uploader.delete_data(authorization_token, catagory, rows)

        return Uploader.get_tanent_access_token(authorization_token)

    def __get_tanent_access_token(last_token):

        conn = http.client.HTTPSConnection("open.feishu.cn")
        payload = json.dumps({
        "app_id": Uploader.app_id,
        "app_secret": Uploader.app_secret
        })
        headers = {
        'Authorization': last_token,
        'Content-Type': 'application/json'
        }
        conn.request("POST", "/open-apis/auth/v3/tenant_access_token/internal", payload, headers)
        response = conn.getresponse()
        resp_bytes = response.read()
        try:
            resp_text = resp_bytes.decode('utf-8')
        except Exception:
            resp_text = resp_bytes.decode(errors='replace')
        try:
            resp_json = json.loads(resp_text)
        except Exception:
            print(f"获取tenant_access_token返回非JSON响应: HTTP {response.status} {response.reason}\n{resp_text}")
            return None
        if resp_json.get('code') == 0:
            token = resp_json.get('tenant_access_token')
            print(f"成功获取新的 tenant_access_token: {token}")
            return f"Bearer {token}"

if __name__ == "__main__":
    
    authorization_token = "Bearer t-g104a6lcW7DWQEFMH2CPFBEDO7UEXMKSAKOSZBJP"
    # 示例用法

    # parse_excel("grouped_data", "20250927_221623")

    # reset_all_sheets()

    # 获取新的 tenant_access_token
    # authorization_token = get_tanent_access_token(authorization_token)