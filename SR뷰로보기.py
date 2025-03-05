import pandas as pd

# 엑셀 파일 불러오기
file_path = "./화성통합 SR내역(25.02).xlsx"
xls = pd.ExcelFile(file_path)

# Sheet1 데이터 로드
df = xls.parse(sheet_name=xls.sheet_names[0])

# HTML에서 개행 반영
df = df.astype(str).applymap(lambda x: x.replace("\n", "<br>"))

# HTML 테이블 변환 (id 추가)
html_table = df.to_html(index=False, escape=False, table_id="data-table")

# HTML 파일 생성 (필터 기능 추가)
html_content = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SR 요청 내역</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 20px;
            padding: 20px;
            background-color: #f4f4f4;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            background-color: white;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }}
        th {{
            background-color: #4CAF50;
            color: white;
        }}
        tr:nth-child(even) {{
            background-color: #f2f2f2;
        }}
        .filter-container {{
            margin-bottom: 15px;
        }}
        .filter-container label {{
            font-weight: bold;
        }}
        .filter-container select {{
            padding: 5px;
            margin-right: 10px;
        }}
    </style>
    <script>
        function filterTable() {{
            var inputs = document.querySelectorAll('.filter-container select');
            var table = document.getElementById("data-table");
            var tr = table.getElementsByTagName("tr");

            for (var i = 1; i < tr.length; i++) {{
                var display = true;
                for (var j = 0; j < inputs.length; j++) {{
                    var column = inputs[j].getAttribute("data-column");
                    var selectedValue = inputs[j].value.toLowerCase();
                    var td = tr[i].getElementsByTagName("td")[column];
                    if (td) {{
                        var cellText = td.textContent.toLowerCase();
                        if (selectedValue !== "" && cellText.indexOf(selectedValue) === -1) {{
                            display = false;
                            break;
                        }}
                    }}
                }}
                tr[i].style.display = display ? "" : "none";
            }}
        }}
    </script>
</head>
<body>
    <h2>SR 요청 내역</h2>

    <div class="filter-container">
"""

# 필터링 가능한 열 추가 (첫 3개 열만 필터 추가)
for i, column in enumerate(df.columns[::]):
    html_content += f"""
        <label for="filter-{i}">{column}</label>
        <select id="filter-{i}" data-column="{i}" onchange="filterTable()">
            <option value="">전체</option>
    """
    unique_values = df[column].dropna().unique()
    for value in sorted(unique_values):
        html_content += f'<option value="{value}">{value}</option>'
    html_content += "</select>"

html_content += "</div>\n"

html_content += html_table

html_content += """
</body>
</html>
"""

# HTML 파일 저장 경로
html_file_path = "./sr_request_summary.html"

# HTML 파일 저장
with open(html_file_path, "w", encoding="utf-8") as f:
    f.write(html_content)

# HTML 파일 경로 반환
html_file_path
