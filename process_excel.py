import pandas as pd
import os

# --- 設定 ---
EXCEL_FILE_PATH = 'data/your_data.xlsx'
HTML_OUTPUT_PATH = 'index.html'
PAGE_TITLE = '宏恩產品查詢'

# --- HTML 模板 ---
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
        /* 我們保留簡潔的基礎樣式 */
        body {{
            font-family: 'Microsoft JhengHei', 'Segoe UI', sans-serif;
            margin: 20px;
        }}
        h1 {{
            text-align: center;
        }}
        /* 讓表格容器寬一點 */
        .container {{
            width: 98%;
            margin: auto;
        }}
        /* DataTables 控制項的間距 */
        .dataTables_wrapper .dataTables_length, .dataTables_wrapper .dataTables_filter {{
            margin-bottom: 1em;
        }}
        /* 確保固定欄位有不透明底色，避免文字透出 */
        table.dataTable tbody tr > .dtfc-fixed-left {{
            background-color: white;
        }}
        table.dataTable thead tr > .dtfc-fixed-left {{
             background-color: #f7f7f7;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>{title}</h1>
        {table}
    </div>

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
            // (我們不再手動設定寬度、不再使用scrollY、不再使用table-layout:fixed)
        }});

        // 加入一個延遲極短的重繪指令，確保初始載入時的對齊萬無一失
        setTimeout(function() {{
            table.columns.adjust().fixedColumns().relayout();
        }}, 10);
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