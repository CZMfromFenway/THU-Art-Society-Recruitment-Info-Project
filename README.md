# 清华大学美术社招新信息化工程项目

## 项目简介

## 文件结构

```
根目录
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

## 使用方法

1. 克隆此仓库到本地：

    ```bash
    git clone

    cd TsinghuaArtClub-Recruitment-InfoSystem
    ```

2. 安装所需的Python库：

    ```bash
    pip install -r requirements.txt
    ```

3. 配置 `config/config.json` 文件，填写相关的API密钥和配置参数。

4. 启动脚本：

    ```bash
    python script/recruitment_data_sync.py --reset
    ```
    该脚本将会自动收集、解析并上传招新数据。按`Ctrl+C`可中断脚本。其中`--reset`参数表示重新收集数据，若脚本运行因故中断，可去掉该参数重新启动，基于已收集的数据继续处理。