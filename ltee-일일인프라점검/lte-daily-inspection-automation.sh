#!/bin/bash
# set -euo pipefail

##############################################################################
# LTE 일일 점검 자동화 스크립트
# 작성자: 박지수
# 최종 작성일: 2025-12-16
# 목적: AWS CloudWatch 지표 수집, 사이트 상태 점검, 엑셀 보고서 생성 및 메일 발송
# 설명:
#   1) CloudWatch API를 통해 CPU, 메모리, 디스크 사용률 수집
#   2) curl을 통해 사이트 상태(HTTP 응답 코드) 점검
#   3) Python(openpyxl)을 통해 엑셀 템플릿 수정
#   4) msmtp를 통해 결과 엑셀 파일 메일 발송
# 사전설정:
#   - AWS CLI, jq, Python(openpyxl), msmtp 설치 및 설정 필요
#   - AWS 자격증명 및 권한 설정 필요 (CloudWatch 지표 조회 권한 포함)
#   - msmtp 설정 파일에 메일 발송 계정 및 SMTP 서버 정보 설정 필요
# 커스텀 필요 사항:
#   - 엑셀 템플릿 파일 경로 및 메일 수신자, 메일 내용 설정 필요
#   - CloudWatch 지표 설정 필요
##############################################################################

### 전역변수
YESTERDAY=$(date -d "yesterday" +"%Y-%m-%d")  # 어제 날짜 (YYYY-MM-DD)
ROOT_TOTAL_GB=50                              # 기본 디스크 전체 용량 (예: 50GB)
DATA_TOTAL_GB=500                             # 추가 디스크 전체 용량 (예: 500GB)
TOTAL_MEM_GIB=16                              # 메모리 전체 용량 (예: 16GiB)

### 커스텀 필요
REGION="ap-northeast-2"
INSTANCE_ID="i-04061836423461dec"
METRICS=(
  "AWS/EC2|CPUUtilization|Name=InstanceId,Value=$INSTANCE_ID"
  "CWAgent|mem_used_percent|Name=InstanceId,Value=$INSTANCE_ID"
  "CWAgent|disk_used_percent|Name=InstanceId,Value=$INSTANCE_ID Name=ImageId,Value=ami-0f073fe0085319698 Name=InstanceType,Value=m5.xlarge Name=device,Value=nvme0n1p1 Name=path,Value=/ Name=fstype,Value=xfs"
  "CWAgent|disk_used_percent|Name=InstanceId,Value=$INSTANCE_ID Name=ImageId,Value=ami-0f073fe0085319698 Name=InstanceType,Value=m5.xlarge Name=device,Value=mapper/lvmVG-lvmLV Name=path,Value=/data Name=fstype,Value=xfs"
)

### ====== 엑셀 템플릿 / 결과 파일 경로 =======
TITLE_DATE=$(date -d "yesterday" +%y.%m.%d)   # 파일명/제목용: xx.xx.xx
OUTDIR="/home/lte/daily-inspection/results"   # 결과물 저장 디렉토리
TEMPLATE="/home/lte/daily-inspection/xx.xx.xx_롯데에너지머티리얼즈_인스턴스_일일사용률_보고서.xlsx"
FILENAME="${TITLE_DATE}_롯데에너지머티리얼즈_인스턴스_일일사용률_보고서.xlsx"
OUTFILE="${OUTDIR}/${FILENAME}"

# PATH 설정 (cron은 PATH가 매우 짧음)
export PATH="/usr/local/bin:/usr/bin:/bin"



# **********************************************************************************************
# === STEP 1. CloudWatch API ===================================================================
# **********************************************************************************************

### CloudWatch API 실행에 필요한 변수
i=0
declare -a DATA_JSON=()                     # 메트릭별 JSON 결과 저장 배열
START="${YESTERDAY}T00:00:00Z"
END="${YESTERDAY}T23:59:59Z"

## METRICS를 순회하며 각 지표 조회
for item in "${METRICS[@]}"; do
  IFS='|' read -r NAMESPACE METRIC_NAME DIMENSION <<< "$item"
  IFS=' ' read -r -a DIM_ARR <<< "$DIMENSION"

  # i. CloudWatch에서 메트릭 통계 가져오기
  if ! DATA_JSON[$i]=$(/usr/local/bin/aws cloudwatch get-metric-statistics \
    --region "$REGION" \
    --namespace "$NAMESPACE" \
    --metric-name "$METRIC_NAME" \
    --dimensions "${DIM_ARR[@]}" \
    --statistics Average \
    --start-time "$START" \
    --end-time "$END" \
    --period 300); then

    # ii. aws 호출 자체가 실패한 경우 호출 실페 메일 발송 후 종료
    # echo "CloudWatch ${METRIC_NAME} 지표 조회 실패"
    msmtp -a bjs bjs@innogrid.com <<EOF
From: "Monitoring" <bjs@innogrid.com>
To: bjs@innogrid.com
Subject: [긴급] CloudWatch ${METRIC_NAME} 지표 수집 실패

현재 인스턴스 ${INSTANCE_ID}의 ${METRIC_NAME} 지표를 CloudWatch에서 조회하지 못했습니다.
스크립트를 확인해 주세요.
EOF
    exit 1
  fi

  # iii. JSON은 받아왔는데 Datapoints 가 비어 있는 경우 실패 메일 발송 후 종료
  DATAPOINT_LEN=$(echo "${DATA_JSON[$i]}" | jq '.Datapoints | length')
  if [ "$DATAPOINT_LEN" -eq 0 ]; then
    echo "Error: No datapoints for mem_used_percent"
    # --- 데이터 포인트가 없는 경우 메일 발송 ---
    msmtp -a bjs bjs@innogrid.com <<EOF
From: "Monitoring" <bjs@innogrid.com>
To: bjs@innogrid.com
Subject: [긴급] CloudWatch ${METRIC_NAME} 지표 데이터 없음

현재 인스턴스 ${INSTANCE_ID}의 ${METRIC_NAME} 지표에 수집된 데이터가 없습니다.
CloudWatch 에이전트 상태 및 지표 설정을 확인해 주세요.
EOF
    exit 1
  fi

  ((i++))
done

# CPU 사용량
CPU_AVG=$(echo "${DATA_JSON[0]}" | jq -r '(.Datapoints // []) | map(.Average) | if length == 0 then 0 else (add / length) end')
CPU_MAX=$(echo "${DATA_JSON[0]}" | jq -r '(.Datapoints // []) | map(.Average) | if length == 0 then 0 else max end')

# 메모리(mem_used_percent)
MEM_AVG_PERCENT=$(echo "${DATA_JSON[1]}" | jq -r '(.Datapoints // []) | map(.Average) | if length == 0 then 0 else (add / length) end')
# 루트 디스크 사용률
ROOT_AVG_PERCENT=$(echo "${DATA_JSON[2]}" | jq -r '(.Datapoints // []) | map(.Average) | if length == 0 then 0 else (add / length) end')

# /data 추가 디스크 사용률
DATA_AVG_PERCENT=$(echo "${DATA_JSON[3]}" | jq -r '(.Datapoints // []) | map(.Average) | if length == 0 then 0 else (add / length) end')



# **********************************************************************************************
# ======= STEP 2. 사이트 상태 체크  ==============================================================
# **********************************************************************************************
HTTP_CODE=$(curl -s -o /dev/null -w "%{http_code}" https://lotteenergymaterials.com/)
if [[ "$HTTP_CODE" =~ ^(200|301|302|403)$ ]]; then
  SITE_STATUS="정상"
else
  SITE_STATUS="비정상"
  # --- 비정상일 때 박지수에게 메일 발송 ---
  msmtp -a bjs bjs@innogrid.com <<EOF
From: "Monitoring" <bjs@innogrid.com>
To: bjs@innogrid.com
Subject: [긴급] 사이트 비정상 - 서버 확인 필요

현재 https://lotteenergymaterials.com/ 사이트 응답 코드가 $HTTP_CODE 로 비정상입니다.
서버 확인이 필요합니다.
EOF
  # --- 스크립트 즉시 종료 ---
  exit 1
fi



# **********************************************************************************************
# ======= STEP 3. Python + openpyxl로 엑셀 수정 ==================================================
# **********************************************************************************************
# 주의! openpyxl 버전 3.0.3 설치 필수
export OUTFILE TITLE_DATE C2_TEXT \
       CPU_AVG CPU_MAX \
       MEM_AVG_PERCENT TOTAL_MEM_GIB \
       ROOT_AVG_PERCENT ROOT_TOTAL_GB \
       DATA_AVG_PERCENT DATA_TOTAL_GB \
       SITE_STATUS

# i. 결과 파일명 생성 및 템플릿 복사
mkdir -p "$OUTDIR"
cp "$TEMPLATE" "$OUTFILE"

# C2 셀용 텍스트
YEAR=$(date -d "yesterday" +%Y)
MONTH=$(date -d "yesterday" +%m)
DAY=$(date -d "yesterday" +%d)
C2_TEXT="일시 : ${YEAR}년 ${MONTH}월 ${DAY}일 00:00~23:59"


python3 << 'PYEOF'
import os
from openpyxl import load_workbook

path = os.environ["OUTFILE"]
title_date = os.environ["TITLE_DATE"]
c2_text = os.environ["C2_TEXT"]

cpu_avg = float(os.environ["CPU_AVG"])
cpu_max = float(os.environ["CPU_MAX"])

mem_avg_percent = float(os.environ["MEM_AVG_PERCENT"])
total_mem_gib = float(os.environ["TOTAL_MEM_GIB"])

root_avg_percent = float(os.environ["ROOT_AVG_PERCENT"])
root_total_gb = float(os.environ["ROOT_TOTAL_GB"])

data_avg_percent = float(os.environ["DATA_AVG_PERCENT"])
data_total_gb = float(os.environ["DATA_TOTAL_GB"])

site_status = os.environ["SITE_STATUS"]

wb = load_workbook(path)
ws = wb.active  # 첫 번째 시트 기준

# 2) C2 셀: "일시 : oooo년 oo월 oo일 00:00~23:59"
ws["C2"] = c2_text

# 3) H5 셀: CPU 최대/평균 (줄바꿈)
#    예: "최대 12.34%\n평균 5.67%"
ws["H5"] = f"최대 {cpu_max:.2f}%\n평균 {cpu_avg:.2f}%"

# 4) I5 셀: 메모리 사용량 "xx.xxG/16G\n(yy.yy%)"
mem_avg_gib = mem_avg_percent / 100.0 * total_mem_gib
ws["I5"] = f"{mem_avg_gib:.2f}G/{total_mem_gib:.0f}G\n({mem_avg_percent:.2f}%)"

# 5) J5 셀: 기본/추가 디스크 평균 사용량
#   "기본: xx.xxGB / 50GB (yy.yy%)\n추가: aa.aaGB / 500GB (bb.bb%)"
root_used_gb = root_avg_percent / 100.0 * root_total_gb
data_used_gb = data_avg_percent / 100.0 * data_total_gb

ws["J5"] = (
    f"기본: {root_used_gb:.2f}GB / {root_total_gb:.0f}GB ({root_avg_percent:.2f}%)\n"
    f"추가: {data_used_gb:.2f}GB / {data_total_gb:.0f}GB ({data_avg_percent:.2f}%)"
)

# 6) K5 셀: 사이트 상태 ("정상" or "이상")
ws["K5"] = site_status

wb.save(path)
PYEOF

echo "생성 완료: $OUTFILE"



# **********************************************************************************************
# ============ STEP 4. 메일 전송 ================================================================
# **********************************************************************************************
BOUNDARY="====BOUNDARY_$(date +%s)==="
ENCODED_FILENAME=$(echo -n "$FILENAME" | iconv -f utf-8 -t euc-kr)
MAIL_SIGN_HTML="$(cat /home/lte/daily-inspection/mail-sign.html)"

# 테스트용: msmtp -a bjs bjs@innogrid.com<<EOF
msmtp -a bjs bjs@innogrid.com yangjeonghwan@lotte.net sunhwa_kim1@lotte.net kjkang1@lotte.net csp@innogrid.com mjh@innogrid.com<<EOF
From: "Innogrid" <bjs@innogrid.com>
To: yangjeonghwan@lotte.net
Cc: sunhwa_kim1@lotte.net, kjkang1@lotte.net, csp@innogrid.com, mjh@innogrid.com
Subject: [롯데에너지머티리얼즈] ${TITLE_DATE} 정보시스템 일일 점검서 전달드립니다
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="$BOUNDARY"

--$BOUNDARY
Content-Type: text/html; charset="UTF-8"
Content-Transfer-Encoding: 8bit

<div>안녕하세요, 이노그리드 박지수입니다.</div>
<div>$TITLE_DATE 롯데에너지머티리얼즈 정보시스템 일일 점검서 전달 드립니다.</div>
<div>첨부 파일 확인 부탁드립니다.</div>
<div>감사합니다.</div>
<div>박지수 드림.</div>


${MAIL_SIGN_HTML}


--$BOUNDARY
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name="$ENCODED_FILENAME"
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="$ENCODED_FILENAME"

$(base64 -w0 "$OUTFILE")
--$BOUNDARY--
EOF