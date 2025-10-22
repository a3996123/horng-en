import pandas as pd
import os

# --- 設定 ---
EXCEL_FILE_PATH = 'data/your_data.xlsx'
HTML_OUTPUT_PATH = 'index.html'
PAGE_TITLE = '宏恩產品查詢'

# --- HTML 模板 (已更新為 '生產批號紀錄' 的 Glassmorphism 樣式) ---
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
        /* --- 樣式基礎 (來自 '生產批號紀錄' index.html) --- */
        body {{ 
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background: linear-gradient(135deg, #a1c4fd 0%, #c2e9fb 100%); 
            min-height: 100vh; 
            box-sizing: border-box; 
        }}
        
        .glass-container {{ 
            background: rgba(255, 255, 255, 0.25); 
            box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.1); 
            backdrop-filter: blur(10px); 
            -webkit-backdrop-filter: blur(10px); 
            border: 1px solid rgba(255, 255, 255, 0.18); 
            border-radius: 16px; 
            padding: 20px;
            width: 98%;
            margin: auto;
            box-sizing: border-box;
        }}

        h1 {{
            text-align: center;
            color: #333; /* 在淺色背景上使用深色文字 */
            text-shadow: 0 1px 2px rgba(255,255,255,0.4);
        }}
        
        /* 4. 懸浮按鈕 (保留) */
        #scrollToTopFab {{
            position: fixed;
            bottom: 30px;
            right: 30px;
            width: 55px;
            height: 55px;
            background-color: #007bff; /* 顏色與新樣式按鈕一致 */
            color: white;
            border: none;
            border-radius: 50%;
            text-align: center;
            font-size: 24px;
            line-height: 55px;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.3);
            cursor: pointer;
            z-index: 1000;
            transition: background-color 0.3s, transform 0.3s;
        }}
        #scrollToTopFab:hover {{
            background-color: #0056b3;
            transform: scale(1.1);
        }}

        /* --- DataTables 玻璃樣式覆蓋 (調整為淺色主題) --- */
        
        /* 控制項 (搜尋、分頁) 的文字顏色 -> 改為深色 */
        .dataTables_wrapper .dataTables_length label,
        .dataTables_wrapper .dataTables_filter label,
        .dataTables_wrapper .dataTables_info,
        .dataTables_wrapper .dataTables_paginate .paginate_button {{
            color: #333 !important;
            font-weight: 500;
        }}
        
        /* 搜尋框 和 長度選擇 (套用新 index 的 input 樣式) */
        .dataTables_wrapper .dataTables_filter input,
        .dataTables_wrapper .dataTables_length select {{
            background: rgba(255, 255, 255, 0.5);
            border: 1px solid rgba(255, 255, 255, 0.3);
            border-radius: 8px;
            color: #333;
            font-size: 14px; /* 調整字體大小以匹配 */
            padding: 8px 10px; /* 調整邊距 */
            margin-left: 5px;
        }}
        .dataTables_wrapper .dataTables_length select option {{
            background: #f9f9f9; /* 下拉選單需要實色背景 */
            color: black;
        }}
        
        /* 表格本體 */
        table.dataTable {{
            width: 100%;
            border-collapse: collapse;
            background: rgba(255, 255, 255, 0.2);
            border-radius: 8px; /* DataTables 會在外層 wrapper 套用... */
            overflow: hidden;
            border: 1px solid rgba(255, 255, 255, 0.3);
        }}

        /* 表頭 (套用新 index 的 th 樣式) */
        table.dataTable thead th {{
            background-color: rgba(255, 255, 255, 0.3);
            color: #333;
            border: 1px solid rgba(255, 255, 255, 0.3);
            padding: 12px;
            text-align: left;
        }}
        
        /* 儲存格 */
        table.dataTable td {{
             border: 1px solid rgba(255, 255, 255, 0.3);
             padding: 10px 12px;
             box-sizing: border-box;
        }}
        
        /* 表格內容行 (保留原本的交錯樣式) */
        table.dataTable tbody tr {{
            color: #333;
        }}
        table.dataTable tbody tr.odd {{
             background: rgba(255, 255, 255, 0.1);
        }}
        table.dataTable tbody tr.even {{
             background: rgba(255, 255, 255, 0.2);
        }}
        
        /* 滑鼠懸停效果 (調整為更亮的白色) */
        table.dataTable tbody tr:hover {{
            background: rgba(255, 255, 255, 0.5) !important;
            color: #000;
        }}
        
        /* 分頁按鈕 */
        .dataTables_wrapper .dataTables_paginate .paginate_button {{
            border: 1px solid rgba(255, 255, 255, 0.3);
            background: rgba(255, 255, 255, 0.2);
            border-radius: 8px;
            margin: 0 3px;
            padding: 6px 10px;
        }}
        .dataTables_wrapper .dataTables_paginate .paginate_button.current,
        .dataTables_wrapper .dataTables_paginate .paginate_button.current:hover {{
            background: #007bff; /* 套用新樣式按鈕色 */
            border-color: #007bff;
            color: #fff !important;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }}
        .dataTables_wrapper .dataTables_paginate .paginate_button:hover {{
            background: rgba(255, 255, 255, 0.5);
            border-color: rgba(255, 255, 255, 0.4);
        }}
        .dataTables_wrapper .dataTables_paginate .paginate_button.disabled {{
            opacity: 0.5;
            background: rgba(255, 255, 255, 0.1);
        }}

        /* * --- 固定欄位 (FixedColumns) 關鍵修復 ---
         * 調整為淺色主題，必須更不透明 + backdrop-filter
        */
        table.dataTable tbody tr > .dtfc-fixed-left,
        table.dataTable tbody tr > .dtfc-fixed-right {{
            background: rgba(255, 255, 255, 0.6); /* 提高不透明度 */
            backdrop-filter: blur(5px);
            -webkit-backdrop-filter: blur(5px);
        }}
        table.dataTable thead tr > .dtfc-fixed-left,
        table.dataTable thead tr > .dtfc-fixed-right {{
             background: rgba(255, 255, 255, 0.7); /* 表頭更不透明 */
             backdrop-filter: blur(5px);
            -webkit-backdrop-filter: blur(5px);
        }}
        
        /* 固定欄位的 Hover 效果 */
         table.dataTable tbody tr:hover > .dtfc-fixed-left,
         table.dataTable tbody tr:hover > .dtfc-fixed-right {{
            background: rgba(255, 255, 255, 0.8) !important;
         }}
    </style>
</head>
<body>
    
    <div class="container glass-container">
        <h1>{title}</h1>
        {table}
    </div>

    <button id="scrollToTopFab" title="回到頂部">&#9650;</button>

    <script>
    $(document).ready( function () {{
        var table = $('#myDataTable').DataTable({{
            "scrollX": true,
            "fixedColumns": {{
                "left": 2
            }},
            "language": {{
                "url": "https://cdn.datatables.net/plug-ins/1.13.6/i18n/zh-HANT.json"
            }}
        }});

        setTimeout(function() {{
            table.columns.adjust().fixedColumns().relayout();
        }}, 10);
        
        // --- 懸浮按鈕點擊事件 (保留) ---
        $('#scrollToTopFab').on('click', function() {{
            $('html, body').animate({{scrollTop: 0}}, 600);
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