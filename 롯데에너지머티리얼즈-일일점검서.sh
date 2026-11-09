#!/bin/bash
set -euo pipefail

#=== 민감 설정 로드 ==============================================
# 인스턴스 ID, 메일 주소, 서명 등은 같은 폴더의 config.env 에서 읽는다 (git 제외)
# 처음 사용할 때: config.env.example 을 config.env 로 복사해 값을 채운다
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CONFIG_FILE="${SCRIPT_DIR}/config.env"
if [ ! -f "$CONFIG_FILE" ]; then
  echo "config.env 파일이 없습니다. config.env.example 을 복사해 작성하세요." >&2
  exit 1
fi
. "$CONFIG_FILE"
MAIL_SIGN_FILE="${MAIL_SIGN_FILE:-${SCRIPT_DIR}/mail-sign.html}"

#=== 설정 =======================================================
REGION="ap-northeast-2"
INSTANCE_ID="${INSTANCE_ID:?config.env 에 INSTANCE_ID 를 설정하세요}"
IMAGE_ID="${IMAGE_ID:?config.env 에 IMAGE_ID 를 설정하세요}"
INSTANCE_TYPE="m5.xlarge"

# 디스크 정보 (d, e 결과용)
ROOT_DEVICE="nvme0n1p1"          # 기본 디스크
ROOT_FSTYPE="xfs"
ROOT_TOTAL_GB=50                 # 기본 디스크 전체 용량 (예: 50GB)

DATA_DEVICE="mapper/lvmVG-lvmLV" # 추가 디스크
DATA_PATH="/data"
DATA_FSTYPE="xfs"
DATA_TOTAL_GB=500                # 추가 디스크 전체 용량 (예: 500GB)

# 메모리 전체 용량 (m5.xlarge = 16GiB)
TOTAL_MEM_GIB=16


### ====== 날짜 계산 (어제) ======
YESTERDAY=$(date -d "yesterday" +"%Y-%m-%d")
YESTERDAY_YMD=$(date -d "yesterday" +%Y-%m-%d)
TITLE_DATE=$(date -d "yesterday" +%y.%m.%d)   # 파일명/제목용: xx.xx.xx
YEAR=$(date -d "yesterday" +%Y)
MONTH=$(date -d "yesterday" +%m)
DAY=$(date -d "yesterday" +%d)

# CloudWatch는 UTC기준 → KST 00:00~23:59를 UTC로 변환
START="${YESTERDAY}T00:00:00Z"
END="${YESTERDAY}T23:59:59Z"


### ====== 엑셀 템플릿 / 결과 파일 경로 =======
TEMPLATE="/home/lte/daily-inspection/xx.xx.xx_롯데에너지머티리얼즈_인스턴스_일일사용률_보고서.xlsx"
OUTDIR="/home/lte/daily-inspection/results"
# C2 셀용 텍스트
C2_TEXT="일시 : ${YEAR}년 ${MONTH}월 ${DAY}일 00:00~23:59"

# PATH 설정 (cron은 PATH가 매우 짧음)
export PATH="/usr/local/bin:/usr/bin:/bin"


### ====== 결과 파일명 생성 및 템플릿 복사 ======
mkdir -p "$OUTDIR"
FILENAME="${TITLE_DATE}_롯데에너지머티리얼즈_인스턴스_일일사용률_보고서.xlsx"
OUTFILE="${OUTDIR}/${FILENAME}"

cp "$TEMPLATE" "$OUTFILE"

#-------------------------------------------------------------
# CPU 사용량
CPU_JSON=$(/usr/local/bin/aws cloudwatch get-metric-statistics \
  --region "$REGION" \
  --namespace "AWS/EC2" \
  --metric-name "CPUUtilization" \
  --dimensions Name=InstanceId,Value="$INSTANCE_ID" \
  --statistics Average \
  --start-time "$START" \
  --end-time "$END" \
  --period 300)

CPU_AVG=$(echo "$CPU_JSON" | jq -r '(.Datapoints // []) | map(.Average) | if length == 0 then 0 else (add / length) end')
CPU_MAX=$(echo "$CPU_JSON" | jq -r '(.Datapoints // []) | map(.Average) | if length == 0 then 0 else max end')


#-------------------------------------------------------------
# 메모리(mem_used_percent)
MEM_JSON=$(/usr/local/bin/aws cloudwatch get-metric-statistics \
  --region "$REGION" \
  --namespace "CWAgent" \
  --metric-name "mem_used_percent" \
  --dimensions Name=InstanceId,Value="$INSTANCE_ID" \
  --statistics Average \
  --start-time "$START" \
  --end-time "$END" \
  --period 300)

# MEM_AVG_PERCENT=$(echo "$MEM_JSON" | jq -r '[.Datapoints[].Average] | add / length')
MEM_AVG_PERCENT=$(echo "$MEM_JSON" | jq -r '(.Datapoints // []) | map(.Average) | if length == 0 then 0 else (add / length) end')


#-------------------------------------------------------------
# 루트 디스크 사용률
ROOT_JSON=$(/usr/local/bin/aws cloudwatch get-metric-statistics \
  --region "$REGION" \
  --namespace "CWAgent" \
  --metric-name "disk_used_percent" \
  --dimensions \
      Name=InstanceId,Value="$INSTANCE_ID" \
      Name=ImageId,Value="$IMAGE_ID" \
      Name=InstanceType,Value="$INSTANCE_TYPE" \
      Name=device,Value="$ROOT_DEVICE" \
      Name=path,Value=/ \
      Name=fstype,Value="$ROOT_FSTYPE" \
  --statistics Average \
  --start-time "$START" \
  --end-time "$END" \
  --period 300)

# ROOT_AVG_PERCENT=$(echo "$ROOT_JSON" | jq -r '[.Datapoints[].Average] | add / length')
ROOT_AVG_PERCENT=$(echo "$ROOT_JSON" | jq -r '(.Datapoints // []) | map(.Average) | if length == 0 then 0 else (add / length) end')


#-------------------------------------------------------------
# /data 추가 디스크 사용률
DATA_JSON=$(/usr/local/bin/aws cloudwatch get-metric-statistics \
  --region $REGION \
  --namespace CWAgent \
  --metric-name disk_used_percent \
  --dimensions \
      Name=InstanceId,Value=$INSTANCE_ID \
      Name=ImageId,Value=$IMAGE_ID \
      Name=InstanceType,Value=$INSTANCE_TYPE \
      Name=device,Value=mapper/lvmVG-lvmLV \
      Name=path,Value=/data \
      Name=fstype,Value=xfs \
  --statistics Average \
  --start-time "$START" \
  --end-time "$END" \
  --period 300)

DATA_AVG_PERCENT=$(echo "$DATA_JSON" | jq -r '(.Datapoints // []) | map(.Average) | if length == 0 then 0 else (add / length) end')


### ====== K5: 사이트 상태 체크 ======
HTTP_CODE=$(curl -s -o /dev/null -w "%{http_code}" https://lotteenergymaterials.com/)
if [[ "$HTTP_CODE" =~ ^(200|301|302|403)$ ]]; then
  SITE_STATUS="정상"
else
  SITE_STATUS="비정상"
  # --- 비정상일 때 담당자에게 메일 발송 ---
  msmtp -a "$MSMTP_ACCOUNT" "$ALERT_TO" <<EOF
From: "Monitoring" <${MAIL_FROM}>
To: ${ALERT_TO}
Subject: [긴급] 사이트 비정상 - 서버 확인 필요

현재 https://lotteenergymaterials.com/ 사이트 응답 코드가 $HTTP_CODE 로 비정상입니다.
서버 확인이 필요합니다.

EOF

  # --- 스크립트 즉시 종료 ---
  exit 1
fi

### ====== Python + openpyxl로 엑셀 수정 ======
export OUTFILE TITLE_DATE C2_TEXT \
       CPU_AVG CPU_MAX \
       MEM_AVG_PERCENT TOTAL_MEM_GIB \
       ROOT_AVG_PERCENT ROOT_TOTAL_GB \
       DATA_AVG_PERCENT DATA_TOTAL_GB \
       SITE_STATUS

# 주의! openpyxl 버전 3.0.3 설치 필수

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

### =========== 메일 전송 ==================
BOUNDARY="====BOUNDARY_$(date +%s)==="

MAIL_SIGN_HTML="$(cat "$MAIL_SIGN_FILE")"
# 테스트용: msmtp -a "$MSMTP_ACCOUNT" "$TEST_MAIL_TO" <<EOF
msmtp -a "$MSMTP_ACCOUNT" "$MAIL_FROM" "$MAIL_TO" ${MAIL_CC//,/ } <<EOF
From: "${MAIL_FROM_NAME}" <${MAIL_FROM}>
To: ${MAIL_TO}
Cc: ${MAIL_CC//,/, }
Subject: [롯데에너지머티리얼즈] ${TITLE_DATE} 정보시스템 일일 점검서 전달드립니다
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="$BOUNDARY"

--$BOUNDARY
Content-Type: text/html; charset="UTF-8"
Content-Transfer-Encoding: 8bit

<div>안녕하세요, ${COMPANY_NAME} ${SENDER_NAME}입니다.</div>
<div>$TITLE_DATE 롯데에너지머티리얼즈 정보시스템 일일 점검서 전달 드립니다.</div>
<div>첨부 파일 확인 부탁드립니다.</div>
<div>감사합니다.</div>
<div>${SENDER_NAME} 드림.</div>
<br>
${MAIL_SIGN_HTML}
--$BOUNDARY
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name="report_$TITLE_DATE.xlsx"
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="$FILENAME"

$(base64 -w0 "$OUTFILE")
--$BOUNDARY--
EOF

echo "메일 전송 완료"

### ===========  ==================