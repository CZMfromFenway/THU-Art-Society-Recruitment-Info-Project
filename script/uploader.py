import pandas as pd
import http.client
import json
import os
import datetime

# 模块说明：用于将本地生成的面试信息表上传到飞书表格，并提供清空/查询表格行数等工具函数。
# 注意：类中部分方法使用双下划线命名（私有），会被 Python 名称改写（name mangling）。

class Uploader:
    """
    Uploader 类封装与飞书表格（Spreadsheet）交互的功能：
    - 上传数据到指定工作表（append 行）
    - 删除指定范围行（清空数据）
    - 获取工作表行数（通过 sheets API）
    - 从本地读取生成的 Excel 并上传
    - 重置（清空）所有小组的表格

    配置项：
    - spreadsheet_token: 每个小组对应的 (spreadsheet_id, sheet_id)
    - app_id / app_secret: 用于获取 tenant_access_token
    """

    # 小组 -> [spreadsheet_id, sheet_id] 映射
    spreadsheet_token = {"书法组": ["IXrnsAVd6h9u2Lt9Tt7cOwEonzg", "01171e"],
                        "国画组": ["EKIosEUJLhc2Z8tvzoecJjCEnWe", "01171e"],
                        "篆刻组": ["EaE0ww3yQiNncbkiqnkcXgjVncd", "81776e"],
                        "西画组": ["Jp1TsCc5QhwaLttc5jvcZwQAnxc", "01171e"],
                        "漫画组": ["Sm0KsWsQohwxrztfAtAcJmDRnSc", "01171e"]}
    
    # 应用凭证（用于内部获取 tenant_access_token）
    app_id = "cli_a86924444cb8d00c"
    app_secret = "ydkzWor6TkJGTa7WZrtHBUfGJ1mNWmn0"

    def __upload_to_feishu(self, authorization_token: str, data: pd.DataFrame, group_name: str):
        """
        私有方法：将 DataFrame 数据追加写入飞书表格指定 sheet。

        参数：
        - authorization_token: Bearer token 字符串
        - data: pandas.DataFrame，要上传的数据（不包含标题行）
        - group_name: 小组名称，用于从 spreadsheet_token 中查表ID
        """
        if group_name not in self.spreadsheet_token:
            print(f"未知小组: {group_name}")
            return

        spreadSheetToken, sheetId = self.spreadsheet_token[group_name]
        path = f"/open-apis/sheets/v2/spreadsheets/{spreadSheetToken}/values_append?insertDataOption=INSERT_ROWS"

        # 准备数据
        values = data.values.tolist()

        # 将其中的nan替换为空字符串

        for r in range(len(values)):
            for c in range(len(values[r])):
                if pd.isna(values[r][c]):
                    values[r][c] = ""

        # 将其中的TimeStamp等数字转换为字符串，避免飞书表格误判
        for r in range(len(values)):
            for c in range(len(values[r])):
                if not isinstance(values[r][c], (float, int)):
                    values[r][c] = str(values[r][c])
        
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

    def __delete_data(self, authorization_token: str, group_name: str, row: int):
        """
        私有方法：删除指定 sheet 中从 startIndex（固定为2）到 endIndex（row）之间的行。

        参数：
        - authorization_token: Bearer token
        - group_name: 小组名称
        - row: 要删除的结束行索引（integer）

        注意：接口使用 DELETE /dimension_range，传入的 sheetId 实为 sheet 的 id（需与 API 文档一致）
        """
        if group_name not in self.spreadsheet_token:
            print(f"未知小组: {group_name}")
            return

        spreadSheetToken, sheedId = self.spreadsheet_token[group_name]
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

    def __get_sheet_rows(self, authorization_token: str, group_name: str) -> int:
        """
        私有方法：获取指定 sheet 的行数或合并单元格信息，用于判断当前已占用的行。

        参数：
        - authorization_token: 带 'Bearer ' 前缀的 tenant_access_token
        - group_name: 小组名称（如 "书法组"）

        返回：
        - 若成功，返回正整数（行数 merges[0].end_row_index）
        - 失败时返回 0
        """
        if group_name not in self.spreadsheet_token:
            print(f"未知小组: {group_name}")
            return 0

        spreadSheetToken, sheedId = self.spreadsheet_token[group_name]

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

    def parse_excel(self, authorization_token: str, dir: str, time_str: str) -> str:
        """
        从指定目录读取按小组命名的面试信息 Excel，并上传到对应的小组表格。

        参数：
        - authorization_token: 用于获取 tenant_access_token 的初始 token（例如管理员 token）
        - dir: 存放本地文件的目录路径
        - time_str: 文件名中的时间戳部分，用于匹配文件名

        返回：
        - 最终使用的 tenant_access_token（带 'Bearer ' 前缀）
        """
        authorization_token = self.__get_tanent_access_token(authorization_token)
        for catagory in self.spreadsheet_token.keys():
            file_path = f"{dir}/{catagory}面试信息_{time_str}.xlsx"
            if os.path.exists(file_path):
                data = pd.read_excel(file_path)
                self.__upload_to_feishu(authorization_token, data, catagory)

        return authorization_token
            
    def reset_all_sheets(self, authorization_token: str) -> str:
        """
        清空所有小组表格中的数据（保留表头）。

        返回新的 tenant_access_token（带 'Bearer ' 前缀）。
        """
        authorization_token = self.__get_tanent_access_token(authorization_token)
        for catagory in self.spreadsheet_token.keys():
            rows = self.__get_sheet_rows(authorization_token, catagory)
            print(f"正在清空 {catagory} 小组数据，共 {rows - 1} 行")
            if rows > 1:
                self.__delete_data(authorization_token, catagory, rows)

        return authorization_token

    def __get_tanent_access_token(self, last_token: str) -> str | None:
        """
        私有方法：使用 app_id/app_secret 请求 tenant_access_token。
        
        参数：
        - last_token: 初始 Authorization（如管理员 token）
        
        返回：
        - 成功：字符串 "Bearer _tenant_access_token_"
        - 失败：None
        """
        conn = http.client.HTTPSConnection("open.feishu.cn")
        payload = json.dumps({
        "app_id": self.app_id,
        "app_secret": self.app_secret
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