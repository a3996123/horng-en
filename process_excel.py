import pandas as pd
import os

# --- 設定 ---
EXCEL_FILE_PATH = 'data/your_data.xlsx'
HTML_OUTPUT_PATH = 'index.html'
PAGE_TITLE = '宏恩產品查詢'

# --- HTML 模板 (已更新為液態玻璃 + 懸浮按鈕) ---
html_template = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.css">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/fixedcolumns/4.3.0/css/fixedColumns.dataTables.min.css">

    <script type="text/javascript" charset="utf8" src="https://code.jquery.com/jquery-3.7.0.js"></script>
    <script type="text/javascript" charset="utf8" src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.js"></script>
    <script type="text/javascript" charset="utf8" src="https://cdn.datatables.net/fixedcolumns/4.3.0/js/dataTables.fixedColumns.min.js"></script>
    
    <style>
        /* --- 液態玻璃 (Glassmorphism) 樣式 --- */
        body {{
            font-family: 'Microsoft JhengHei', 'Segoe UI', sans-serif;
            margin: 0;
            padding: 40px 20px; /* 讓容器周圍有空間 */
            color: #333;
            overflow-x: hidden; /* 防止 body 水平滾動 */
        }}
        
        /* 1. 背景 */
        .bg {{
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: linear-gradient(120deg, #89f7fe 0%, #66a6ff 100%);
            z-index: -1;
        }}
        
        /* 2. 玻璃容器 */
        .container {{
            width: 98%;
            margin: auto;
            background: rgba(255, 255, 255, 0.15); /* 玻璃質感 */
            backdrop-filter: blur(15px);
            -webkit-backdrop-filter: blur(15px); /* Safari 支援 */
            border-radius: 20px;
            border: 1px solid rgba(255, 255, 255, 0.2);
            box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.37);
            padding: 20px 25px;
            box-sizing: border-box;
        }}

        /* 3. 標題 */
        h1 {{
            text-align: center;
            color: #fff;
            text-shadow: 0 2px 4px rgba(0,0,0,0.2);
        }}
        
        /* 4. 懸浮按鈕 */
        #scrollToTopFab {{
            position: fixed;
            bottom: 30px;
            right: 30px;
            width: 55px;
            height: 55px;
            background-color: #007bff;
            color: white;
            border: none;
            border-radius: 50%;
            text-align: center;
            font-size: 24px;
            line-height: 55px; /* 垂直置中箭頭 */
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.3);
            cursor: pointer;
            z-index: 1000;
            transition: background-color 0.3s, transform 0.3s;
        }}
        #scrollToTopFab:hover {{
            background-color: #0056b3;
            transform: scale(1.1);
        }}

        /* --- DataTables 玻璃樣式覆蓋 --- */
        
        /* 控制項 (搜尋、分頁) 的文字顏色 */
        .dataTables_wrapper .dataTables_length label,
        .dataTables_wrapper .dataTables_filter label,
        .dataTables_wrapper .dataTables_info,
        .dataTables_wrapper .dataTables_paginate .paginate_button {{
            color: #eee !important;
        }}
        
        /* 搜尋框 和 長度選擇 */
        .dataTables_wrapper .dataTables_filter input,
        .dataTables_wrapper .dataTables_length select {{
            background: rgba(255, 255, 255, 0.3);
            border: 1px solid rgba(255, 255, 255, 0.5);
            color: #fff;
            border-radius: 5px;
            padding: 5px;
        }}
        .dataTables_wrapper .dataTables_length select option {{
            background: #334; /* 下拉選單需要實色背景 */
            color: white;
        }}
        
        /* 表格本體 */
        table.dataTable {{
            border: 1px solid rgba(255, 255, 255, 0.2);
            border-collapse: collapse; /* 移除雙邊框 */
        }}

        /* 表頭 */
        table.dataTable thead th, table.dataTable thead td {{
            background: rgba(255, 255, 255, 0.2);
            color: #fff;
            border-bottom: 1px solid rgba(255, 255, 255, 0.3);
        }}
        
        /* 表格內容行 */
        table.dataTable tbody tr {{
            color: #333; /* 保持深的文字顏色以供閱讀 */
        }}
        table.dataTable tbody tr.odd {{
             background: rgba(255, 255, 255, 0.1);
        }}
        table.dataTable tbody tr.even {{
             background: rgba(255, 255, 255, 0.2);
        }}
        
        /* 儲存格邊框 */
        table.dataTable td, table.dataTable th {{
             border-right: 1px solid rgba(255, 255, 255, 0.15);
             box-sizing: border-box;
        }}
        
        /* 滑鼠懸停效果 */
        table.dataTable tbody tr:hover {{
            background: rgba(255, 255, 255, 0.4) !important; /* 提高對比度 */
            color: #000;
        }}
        
        /* 分頁按鈕 */
        .dataTables_wrapper .dataTables_paginate .paginate_button {{
            border: 1px solid rgba(255, 255, 255, 0.3);
            background: rgba(255, 255, 255, 0.1);
            border-radius: 5px;
            margin: 0 3px;
        }}
        .dataTables_wrapper .dataTables_paginate .paginate_button.current,
        .dataTables_wrapper .dataTables_paginate .paginate_button.current:hover {{
            background: #007bff;
            border-color: #007bff;
            color: #fff !important;
        }}
        .dataTables_wrapper .dataTables_paginate .paginate_button:hover {{
            background: rgba(255, 255, 255, 0.3);
            border-color: rgba(255, 255, 255, 0.4);
        }}
        .dataTables_wrapper .dataTables_paginate .paginate_button.disabled {{
            opacity: 0.4;
        }}

        /* * --- 固定欄位 (FixedColumns) 關鍵修復 ---
         * 必須設置一個不透明度更高的背景 + backdrop-filter
         * 否則滾動時，後面的文字會穿透固定欄位。
        */
        table.dataTable tbody tr > .dtfc-fixed-left,
        table.dataTable tbody tr > .dtfc-fixed-right {{
            background: rgba(245, 245, 245, 0.4); /* 提高不透明度 */
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
        }}
        table.dataTable thead tr > .dtfc-fixed-left,
        table.dataTable thead tr > .dtfc-fixed-right {{
             background: rgba(240, 240, 240, 0.5); /* 表頭更不透明 */
             backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
        }}
        
        /* 固定欄位的 Hover 效果 */
         table.dataTable tbody tr:hover > .dtfc-fixed-left,
         table.dataTable tbody tr:hover > .dtfc-fixed-right {{
            background: rgba(255, 255, 255, 0.6) !important;
         }}
    </style>
</head>
<body>
    
    <div class="bg"></div>

    <div class="container">
        <h1>{title}</h1>
        {table}
    </div>

    <button id="scrollToTopFab" title="回到頂部">&#9650;</button>

    <script>
    $(document).ready( function () {{
        var table = $('#myDataTable').DataTable({{
            // --- 最簡化的核心設定 ---
            "scrollX": true,       // 啟用水平捲動
            "fixedColumns": {{
                "left": 2          // 固定左邊 2 欄
            }},
            
            // --- 其他輔助設定 ---
            "language": {{
                "url": "https://cdn.datatables.net/plug-ins/1.13.6/i18n/zh-HANT.json"
            }}
        }});

        // 加入一個延遲極短的重繪指令，確保初始載入時的對齊萬無一失
        setTimeout(function() {{
            table.columns.adjust().fixedColumns().relayout();
        }}, 10);
        
        // --- 懸浮按鈕點擊事件 ---
        $('#scrollToTopFab').on('click', function() {{
            $('html, body').animate({{scrollTop: 0}}, 600); // 600 毫秒平滑滾動
        }});
        
    }} );
    </script>
</body>
</html>
"""

def create_html_from_excel():
    """
    讀取 Excel 檔案並生成一個簡潔、帶有核心互動功能的 HTML 網頁。
    """
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"錯誤: 找不到 Excel 檔案 at '{EXCEL_FILE_PATH}'")
        return

    try:
        # 讀取 Excel 檔案，並將第一行作為欄位標題 (header=0)
        df = pd.read_excel(EXCEL_FILE_PATH, header=0)
        df = df.fillna('')
        
        # 為了讓 DataTables 能正確初始化，我們為 table 加上 id
        # Pandas 的 to_html 預設會產生 class="dataframe"，我們可以用它來加基礎樣式
        html_table = df.to_html(escape=False, index=False, table_id='myDataTable')
        
        # 將生成的表格填入 HTML 模板
        final_html = html_template.format(title=PAGE_TITLE, table=html_table)
        
        with open(HTML_OUTPUT_PATH, 'w', encoding='utf-8') as f:
            f.write(final_html)
            
        print(f"成功！ 已生成帶有搜尋、排序、固定欄位功能的 HTML 檔案。")

    except Exception as e:
        print(f"處理過程中發生錯誤: {e}")

if __name__ == '__main__':
    create_html_from_excel()