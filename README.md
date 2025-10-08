# 清华大学美术社招新信息化工程项目

## 项目简介

本项目旨在简化和自动化清华大学美术社的招新信息收集与管理流程。通过集成问卷星的数据收集功能和飞书的协作平台，实现从问卷数据的自动收集、解析到上传至飞书文档的全流程自动化，提升招新效率，减少人工操作。

其目标应用场景是：在招新期间，于美社“服务器”启动脚本后，脚本能够全自动定时自动从问卷星收集最新的招新数据，解析并分类这些数据，然后将其上传到飞书文档中，供面试官和相关人员查看和管理。

本项目默认用于Windows11环境，其他操作系统未测试。

## 文件结构

```
项目根目录
│  .gitignore                                  （git忽略文件）
│  README.md                                    (此文件)
│  requirements.txt                            （Python依赖库列表）
│
├─config
│      config.json                             （配置文件）
│
├─grouped_data
│      {组别名}面试信息_{yyyyMMdd_hhmmss}.xlsx  （分组后的面试信息）
│
├─raw_data
│      问卷数据.xlsx                            （从问卷星导出的原始数据）
│
└─script
    │  collect_raw_data.py                     （收集原始数据脚本）
    │  parse_raw_data.py                       （解析原始数据脚本）
    │  recruitment_data_sync.py                （同步招新数据总脚本）
    │  uploader.py                             （上传数据脚本）
    │
    └─test                                     （测试脚本或废稿）    
            feishu.py
            test1.py
            问卷收集_带数据分组_已废弃.py
```
其中`raw_data`和`grouped_data`文件夹用于存放数据文件，缺失宜自建。

## 使用方法

0. 环境准备

    - Python 3.7及以上版本（开发时使用的版本为3.13.2）
    - 安装Git

1. 克隆此仓库到本地

    在目标目录下打开Git Bash，运行以下命令：
    ```bash
    git clone https://github.com/CZMfromFenway/THU-Art-Society-Recruitment-Info-Project.git
    ```
    或直接从Github页面下载ZIP文件并解压。

2. 安装所需的Python库

    ```bash
    pip install -r requirements.txt
    ```

    注意：请确保在项目根目录下运行此命令，Windows用户（以Win11为例）可在项目根目录文件夹中右键->显示更多选项->在终端中打开。

3. 配置 `config/config.json` 文件，填写相关的API密钥和配置参数。

    `config.json` 文件结构如下：
    ```json
    {
        "wjx_url": "https://www.wjx.cn/joinnew/JoinNew.aspx?activity=12345678",
        "wjx_cookie": "your_cookie_here",
        "raw_data_file": "raw_data/问卷数据.xlsx",
        "grouped_data_dir": "grouped_data",
        "feishu_token": "your_feishu_token_here",
        "period": 500
    }
    ```
    - `wjx_url`: 问卷星问卷所导出excel表格的URL地址。具体获取方法如下。
    
        - 登录问卷星->找到相应问卷点击右下角“分析&下载”->查看下载答卷->下载答卷数据->右键“按序号下载Excel”->复制链接
        
        - 将剪贴板中的链接复制到引号中即可。

    - `wjx_cookie`: 访问问卷星所需的Cookie信息。具体获取方法如下。

        - 登陆问卷星（停在登录后的界面）->按F12打开开发者工具->选择“网络”选项卡->刷新页面->在左侧请求列表中选择第一个请求（通常是`myquestionnaires.aspx`）->在右侧选择`标头`选项卡->向下滚动找到`Cookie`字段->复制其内容

        - 将剪贴板中的链接复制到引号中即可。

    - `raw_data_file`: 原始数据文件路径，通常为 `raw_data/问卷数据.xlsx`。无需更改，但要确保根目录下有此文件夹。
    - `grouped_data_dir`: 分组后的数据存放目录，通常为 `grouped_data`。无需更改，但要确保根目录下有此文件夹。
    - `feishu_token`: 飞书API的访问令牌，用于上传数据到飞书。

        - 登录飞书开放平台->点击左上角`开发文档`->点击右上角`API调试台`->找到请求头中的`Authorization`字段->确保此处是`tenant_access_token`->复制其内容

        - 内容样例
            ```
            Bearer t-g104a8adU24XEMINHF364TTF4LOQ4FDQH7HUCAF3
            ```
        - 将剪贴板中的链接复制到引号中即可。

    - `period`: 上传数据的时间间隔，单位为秒，可按需要修改。

4. 启动脚本

    在项目根目录下打开终端，运行以下命令启动脚本：
    ```bash
    python script/recruitment_data_sync.py --reset
    ```
    该脚本将会自动收集、解析并上传招新数据。按`Ctrl+C`可中断脚本。其中`--reset`参数表示重新收集数据，若脚本运行因故中断，可去掉该参数重新启动，基于已收集的数据继续处理。

    **警告：** 重置数据意味着清除飞书文档中的已有数据，清除本地存储的问卷原始数据和分类后的数据，若面试人员已在飞书文档中编辑了面试信息，请勿使用该参数。

程紫陌 20251008