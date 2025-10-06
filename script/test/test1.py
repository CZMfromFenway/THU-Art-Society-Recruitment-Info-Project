import http.client
import json

conn = http.client.HTTPSConnection("open.feishu.cn")
payload = json.dumps({
   "valueRange": {
      "range": "01171e!A2:O3",
      "values": [
         [
             
"2025-09-18 12:08:46",	"张堃",	"为先书院",	"51",	"18214471674",	"_62079913zhang",	2,	0,	0,	0,	0,	1,	0,		"技能：平面设计；有兴趣加入宣传小组"

         ],
        [
             
"2025-09-18 12:08:46",	"张堃",	"为先书院",	"51",	"18214471674",	"_62079913zhang",	2,	0,	0,	0,	0,	1,	0,		"技能：平面设计；有兴趣加入宣传小组"

         ]
      ]
   }
})
headers = {
   'Authorization': 'Bearer u-f_iKS5zuJ0_9Dz5iaAbAvS1lgLcl0lwphgG0g5202I7q',
   'Content-Type': 'application/json'
}
conn.request("POST", "/open-apis/sheets/v2/spreadsheets/IXrnsAVd6h9u2Lt9Tt7cOwEonzg/values_append?insertDataOption=INSERT_ROWS", payload, headers)
res = conn.getresponse()
data = res.read()
print(data.decode("utf-8"))