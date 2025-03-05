import pandas as pd
import re

# 파일 경로
file_path = "./화성시 문의이력 25년 2월.xlsx"

# 엑셀 파일 읽기
df = pd.read_excel(file_path, sheet_name="Sheet1", dtype=str, skiprows=2)  # 3번째 줄을 컬럼명으로 사용

# 함수 정의
def process_m_column(m_value):
    if pd.isna(m_value):
        return "", ""
    
    # &nbsp; → 공백 변환
    m_value = m_value.replace("&nbsp;", " ")

    # "사용자의 성함: "과 "연락처: " 사이의 값을 추출하여 G열로 이동
    name_match = re.search(r"사용자의 성함:\s*(.*?)\s*연락처:", m_value)
    user_name = name_match.group(1) if name_match else ""

    # "문의내용: " 뒤의 내용만 남김
    inquiry_match = re.search(r"문의내용:\s*(.*)", m_value)
    inquiry_text = inquiry_match.group(1) if inquiry_match else ""

    return user_name, inquiry_text

# M열을 변환하고 G열 업데이트
df["요청자"], df["요청내용"] = zip(*df["요청내용"].map(process_m_column))

# 수정된 내용을 sheet2에 저장
with pd.ExcelWriter(file_path, engine="openpyxl", mode="a") as writer:
    df.to_excel(writer, sheet_name="sheet2", index=False)