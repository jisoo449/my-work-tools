@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

::############################################################################
:: LTE 일일 점검 자동화 스크립트 (Windows 변환본)
:: 원작성자: 박지수
:: 변환일: 2025-12-16
:: 사전 설치 필요:
::   - AWS CLI (https://aws.amazon.com/cli/)
::   - AWS Configure 설정
::   - jq.exe  (https://jqlang.github.io/jq/download/)
::   - Python + openpyxl  (pip install openpyxl==3.0.3)
::   - curl (Windows 10 이상 기본 내장)
::   - Microsoft Outlook (설치 및 계정 로그인 필요)
::   - powershell
::############################################################################


:: ============================================================
:: [사전설정] 커스텀 필요 항목
:: ============================================================
set REGION=ap-northeast-2
set INSTANCE_ID=i-04061836423461dec
set ROOT_TOTAL_GB=50
set DATA_TOTAL_GB=500
set TOTAL_MEM_GIB=16

:: jq 경로 (jq.exe를 PATH에 넣었으면 그냥 jq 로 써도 됨)
set JQ=C:\Users\INNOGRID\AppData\Local\Microsoft\WinGet\Packages\jqlang.jq_Microsoft.Winget.Source_8wekyb3d8bbwe\jq.exe

:: 결과물 저장 디렉토리 및 템플릿 경로
set OUTDIR=C:\Users\INNOGRID\개인용 폴더\my-work-tools\ltee-일일인프라점검\results
set TEMPLATE=C:\Users\INNOGRID\개인용 폴더\my-work-tools\ltee-일일인프라점검\xx.xx.xx_롯데에너지머티리얼즈_인스턴스_일일사용률_보고서.xlsx

:: 메일 서명 경로
set MAIL_FROM=bjs@innogrid.com
set MAIL_SIGN=C:\Users\INNOGRID\개인용 폴더\my-work-tools\ltee-일일인프라점검\mail-sign.html
set MAIL_PS_SCRIPT=C:\Users\INNOGRID\개인용 폴더\my-work-tools\ltee-일일인프라점검\send-mail.ps1

:: ============================================================
:: 날짜 계산 (어제)
:: ============================================================
for /f "tokens=*" %%d in ('powershell -NoProfile -Command "(Get-Date).AddDays(-1).ToString('yyyy-MM-dd')"') do set YESTERDAY=%%d
for /f "tokens=*" %%d in ('powershell -NoProfile -Command "(Get-Date).AddDays(-1).ToString('yy.MM.dd')"') do set TITLE_DATE=%%d
for /f "tokens=*" %%d in ('powershell -NoProfile -Command "(Get-Date).AddDays(-1).ToString('yyyy')"') do set YEAR=%%d
for /f "tokens=*" %%d in ('powershell -NoProfile -Command "(Get-Date).AddDays(-1).ToString('MM')"') do set MONTH=%%d
for /f "tokens=*" %%d in ('powershell -NoProfile -Command "(Get-Date).AddDays(-1).ToString('dd')"') do set DAY=%%d

set START=%YESTERDAY%T00:00:00Z
set END=%YESTERDAY%T23:59:59Z
set FILENAME=%TITLE_DATE%_롯데에너지머티리얼즈_인스턴스_일일사용률_보고서.xlsx
set OUTFILE=%OUTDIR%\%FILENAME%
set C2_TEXT=일시 : %YEAR%년 %MONTH%월 %DAY%일 00:00~23:59


:: ============================================================
:: STEP 1. CloudWatch 지표 수집
:: ============================================================
echo [STEP 1] CloudWatch 지표 수집 중...

:: --- CPU ---
aws cloudwatch get-metric-statistics ^
  --region %REGION% ^
  --namespace AWS/EC2 ^
  --metric-name CPUUtilization ^
  --dimensions Name=InstanceId,Value=%INSTANCE_ID% ^
  --statistics Average ^
  --start-time %START% --end-time %END% ^
  --period 300 > "%TEMP%\cw_cpu.json" 2>&1
if errorlevel 1 goto :MAIL_CW_FAIL_CPUUtilization

for /f "tokens=*" %%v in ('%JQ% -r "(.Datapoints // []) | map(.Average) | if length == 0 then 0 else (add / length) end" "%TEMP%\cw_cpu.json"') do set CPU_AVG=%%v
for /f "tokens=*" %%v in ('%JQ% -r "(.Datapoints // []) | map(.Average) | if length == 0 then 0 else max end" "%TEMP%\cw_cpu.json"') do set CPU_MAX=%%v
for /f "tokens=*" %%v in ('%JQ% ".Datapoints | length" "%TEMP%\cw_cpu.json"') do set CPU_LEN=%%v
if "%CPU_LEN%"=="0" goto :MAIL_CW_NODATA_CPUUtilization

:: --- 메모리 ---
aws cloudwatch get-metric-statistics ^
  --region %REGION% ^
  --namespace CWAgent ^
  --metric-name mem_used_percent ^
  --dimensions Name=InstanceId,Value=%INSTANCE_ID% ^
  --statistics Average ^
  --start-time %START% --end-time %END% ^
  --period 300 > "%TEMP%\cw_mem.json" 2>&1
if errorlevel 1 goto :MAIL_CW_FAIL_mem_used_percent

for /f "tokens=*" %%v in ('%JQ% -r "(.Datapoints // []) | map(.Average) | if length == 0 then 0 else (add / length) end" "%TEMP%\cw_mem.json"') do set MEM_AVG_PERCENT=%%v
for /f "tokens=*" %%v in ('%JQ% ".Datapoints | length" "%TEMP%\cw_mem.json"') do set MEM_LEN=%%v
if "%MEM_LEN%"=="0" goto :MAIL_CW_NODATA_mem_used_percent

:: --- 루트 디스크 ---
aws cloudwatch get-metric-statistics ^
  --region %REGION% ^
  --namespace CWAgent ^
  --metric-name disk_used_percent ^
  --dimensions ^
    Name=InstanceId,Value=%INSTANCE_ID% ^
    Name=ImageId,Value=ami-0f073fe0085319698 ^
    Name=InstanceType,Value=m5.xlarge ^
    Name=device,Value=nvme0n1p1 ^
    Name=path,Value=/ ^
    Name=fstype,Value=xfs ^
  --statistics Average ^
  --start-time %START% --end-time %END% ^
  --period 300 > "%TEMP%\cw_root.json" 2>&1
if errorlevel 1 goto :MAIL_CW_FAIL_disk_used_percent_root

for /f "tokens=*" %%v in ('%JQ% -r "(.Datapoints // []) | map(.Average) | if length == 0 then 0 else (add / length) end" "%TEMP%\cw_root.json"') do set ROOT_AVG_PERCENT=%%v
for /f "tokens=*" %%v in ('%JQ% ".Datapoints | length" "%TEMP%\cw_root.json"') do set ROOT_LEN=%%v
if "%ROOT_LEN%"=="0" goto :MAIL_CW_NODATA_disk_used_percent_root

:: --- /data 추가 디스크 ---
aws cloudwatch get-metric-statistics ^
  --region %REGION% ^
  --namespace CWAgent ^
  --metric-name disk_used_percent ^
  --dimensions ^
    Name=InstanceId,Value=%INSTANCE_ID% ^
    Name=ImageId,Value=ami-0f073fe0085319698 ^
    Name=InstanceType,Value=m5.xlarge ^
    Name=device,Value=mapper/lvmVG-lvmLV ^
    Name=path,Value=/data ^
    Name=fstype,Value=xfs ^
  --statistics Average ^
  --start-time %START% --end-time %END% ^
  --period 300 > "%TEMP%\cw_data.json" 2>&1
if errorlevel 1 goto :MAIL_CW_FAIL_disk_used_percent_data

for /f "tokens=*" %%v in ('%JQ% -r "(.Datapoints // []) | map(.Average) | if length == 0 then 0 else (add / length) end" "%TEMP%\cw_data.json"') do set DATA_AVG_PERCENT=%%v
for /f "tokens=*" %%v in ('%JQ% ".Datapoints | length" "%TEMP%\cw_data.json"') do set DATA_LEN=%%v
if "%DATA_LEN%"=="0" goto :MAIL_CW_NODATA_disk_used_percent_data

echo [STEP 1] 완료 - CPU_AVG=%CPU_AVG% MEM=%MEM_AVG_PERCENT% ROOT=%ROOT_AVG_PERCENT% DATA=%DATA_AVG_PERCENT%


:: ============================================================
:: STEP 2. 사이트 상태 체크
:: ============================================================
echo [STEP 2] 사이트 상태 점검 중...

for /f "tokens=*" %%c in ('curl -s -o NUL -w "%%{http_code}" https://lotteenergymaterials.com/') do set HTTP_CODE=%%c

set SITE_STATUS=비정상
if "%HTTP_CODE%"=="200" ( 
  set SITE_STATUS=정상 
) else ( 
    if "%HTTP_CODE%"=="301" ( 
      set SITE_STATUS=정상
    ) else (
        if "%HTTP_CODE%"=="302" (
          set SITE_STATUS=정상
        )
      )
    )

if "%SITE_STATUS%"=="비정상" (
  echo [STEP 2] 사이트 비정상 - HTTP %HTTP_CODE%
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ol = New-Object -ComObject Outlook.Application;" ^
    "$mail = $ol.CreateItem(0);" ^
    "$mail.To = '%MAIL_FROM%';" ^
    "$mail.Subject = '[긴급] 사이트 비정상 - 서버 확인 필요';" ^
    "$mail.Body = '현재 https://lotteenergymaterials.com/ 사이트 응답 코드가 %HTTP_CODE% 로 비정상입니다. 서버 확인이 필요합니다.';" ^
    "$mail.Send()"
  exit /b 1
)
echo [STEP 2] 완료 - 사이트 %SITE_STATUS% (HTTP %HTTP_CODE%)


:: ============================================================
:: STEP 3. Python + openpyxl 엑셀 수정
:: ============================================================
echo [STEP 3] 엑셀 보고서 생성 중...

if not exist "%OUTDIR%" mkdir "%OUTDIR%"
copy /Y "%TEMPLATE%" "%OUTFILE%" > nul

python -c ^
"import os, shutil; from openpyxl import load_workbook; ^
cpu_avg=float('%CPU_AVG%'); cpu_max=float('%CPU_MAX%'); ^
mem_avg=float('%MEM_AVG_PERCENT%'); tmem=float('%TOTAL_MEM_GIB%'); ^
root_avg=float('%ROOT_AVG_PERCENT%'); troot=float('%ROOT_TOTAL_GB%'); ^
data_avg=float('%DATA_AVG_PERCENT%'); tdata=float('%DATA_TOTAL_GB%'); ^
wb=load_workbook('%OUTFILE:\=/%'); ws=wb.active; ^
ws['C2']='%C2_TEXT%'; ^
ws['H5']=f'최대 {cpu_max:.2f}%%\n평균 {cpu_avg:.2f}%%'; ^
ws['I5']=f'{mem_avg/100*tmem:.2f}G/{tmem:.0f}G\n({mem_avg:.2f}%%)'; ^
ws['J5']=f'기본: {root_avg/100*troot:.2f}GB / {troot:.0f}GB ({root_avg:.2f}%%)\n추가: {data_avg/100*tdata:.2f}GB / {tdata:.0f}GB ({data_avg:.2f}%%)'; ^
ws['K5']='%SITE_STATUS%'; ^
wb.save('%OUTFILE:\=/%')"

if errorlevel 1 (
  echo [STEP 3] Python 엑셀 생성 실패
  exit /b 1
)
echo [STEP 3] 완료 - %OUTFILE%


:: ============================================================
:: STEP 4. 메일 발송 (Outlook COM)
:: ============================================================
echo [STEP 4] 메일 발송 중...

@REM powershell -NoProfile -ExecutionPolicy Bypass -Command ^
@REM "$signHtml = if (Test-Path '%MAIL_SIGN%') { Get-Content '%MAIL_SIGN%' -Raw } else { '' };" ^
@REM "$body = '<div>안녕하세요, 이노그리드 박지수입니다.</div>' +" ^
@REM "        '<div>%TITLE_DATE% 롯데에너지머티리얼즈 정보시스템 일일 점검서 전달 드립니다.</div>' +" ^
@REM "        '<div>첨부 파일 확인 부탁드립니다.</div><div>감사합니다.</div><div>박지수 드림.</div>' +" ^
@REM "        $signHtml;" ^
@REM "$ol = New-Object -ComObject Outlook.Application;" ^
@REM "$mail = $ol.CreateItem(0);" ^
@REM "$mail.To = '%MAIL_FROM%';" ^
@REM "$mail.Subject = '[롯데에너지머티리얼즈] %TITLE_DATE% 정보시스템 일일 점검서 전달드립니다';" ^
@REM "$mail.HTMLBody = $body;" ^
@REM "$mail.Attachments.Add('%OUTFILE%') | Out-Null;" ^
@REM "$mail.Send()"

echo TITLE_DATE=%TITLE_DATE%
echo MAIL_SIGN=%MAIL_SIGN%
echo OUTFILE=%OUTFILE%
echo MAIL_FROM=%MAIL_FROM%

powershell -NoProfile -ExecutionPolicy Bypass -File "%MAIL_PS_SCRIPT%" -MailSign "%MAIL_SIGN%" -TitleDate "%TITLE_DATE%" -OutFile "%OUTFILE%" -MailFrom "%MAIL_FROM%"

if errorlevel 1 (
  echo [STEP 4] 메일 발송 실패
  exit /b 1
)
echo [STEP 4] 메일 발송 완료
echo.
echo === 스크립트 정상 완료 ===
exit /b 0


:: ============================================================
:: 오류 처리 라벨
:: ============================================================
:MAIL_CW_FAIL_CPUUtilization
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$ol = New-Object -ComObject Outlook.Application; $mail = $ol.CreateItem(0);" ^
"$mail.To = '%MAIL_FROM%';" ^
"$mail.Subject = '[긴급] CloudWatch CPUUtilization 지표 수집 실패';" ^
"$mail.Body = '인스턴스 %INSTANCE_ID%의 CPUUtilization 지표를 CloudWatch에서 조회하지 못했습니다.';" ^
"$mail.Send()"
exit /b 1C

:MAIL_CW_NODATA_CPUUtilization
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$ol = New-Object -ComObject Outlook.Application; $mail = $ol.CreateItem(0);" ^
"$mail.To = '%MAIL_FROM%';" ^
"$mail.Subject = '[긴급] CloudWatch CPUUtilization 지표 데이터 없음';" ^
"$mail.Body = '인스턴스 %INSTANCE_ID%의 CPUUtilization 지표에 수집된 데이터가 없습니다.';" ^
"$mail.Send()"
exit /b 1

:MAIL_CW_FAIL_mem_used_percent
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$ol = New-Object -ComObject Outlook.Application; $mail = $ol.CreateItem(0);" ^
"$mail.To = '%MAIL_FROM%';" ^
"$mail.Subject = '[긴급] CloudWatch mem_used_percent 지표 수집 실패';" ^
"$mail.Body = '인스턴스 %INSTANCE_ID%의 mem_used_percent 지표를 CloudWatch에서 조회하지 못했습니다.';" ^
"$mail.Send()"
exit /b 1

:MAIL_CW_NODATA_mem_used_percent
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$ol = New-Object -ComObject Outlook.Application; $mail = $ol.CreateItem(0);" ^
"$mail.To = '%MAIL_FROM%';" ^
"$mail.Subject = '[긴급] CloudWatch mem_used_percent 지표 데이터 없음';" ^
"$mail.Body = '인스턴스 %INSTANCE_ID%의 mem_used_percent 지표에 수집된 데이터가 없습니다.';" ^
"$mail.Send()"
exit /b 1

:MAIL_CW_FAIL_disk_used_percent_root
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$ol = New-Object -ComObject Outlook.Application; $mail = $ol.CreateItem(0);" ^
"$mail.To = '%MAIL_FROM%';" ^
"$mail.Subject = '[긴급] CloudWatch disk_used_percent(root) 지표 수집 실패';" ^
"$mail.Body = '인스턴스 %INSTANCE_ID%의 disk_used_percent(/) 지표를 CloudWatch에서 조회하지 못했습니다.';" ^
"$mail.Send()"
exit /b 1

:MAIL_CW_NODATA_disk_used_percent_root
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$ol = New-Object -ComObject Outlook.Application; $mail = $ol.CreateItem(0);" ^
"$mail.To = '%MAIL_FROM%';" ^
"$mail.Subject = '[긴급] CloudWatch disk_used_percent(root) 지표 데이터 없음';" ^
"$mail.Body = '인스턴스 %INSTANCE_ID%의 disk_used_percent(/) 지표에 수집된 데이터가 없습니다.';" ^
"$mail.Send()"
exit /b 1

:MAIL_CW_FAIL_disk_used_percent_data
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$ol = New-Object -ComObject Outlook.Application; $mail = $ol.CreateItem(0);" ^
"$mail.To = '%MAIL_FROM%';" ^
"$mail.Subject = '[긴급] CloudWatch disk_used_percent(/data) 지표 수집 실패';" ^
"$mail.Body = '인스턴스 %INSTANCE_ID%의 disk_used_percent(/data) 지표를 CloudWatch에서 조회하지 못했습니다.';" ^
"$mail.Send()"
exit /b 1

:MAIL_CW_NODATA_disk_used_percent_data
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$ol = New-Object -ComObject Outlook.Application; $mail = $ol.CreateItem(0);" ^
"$mail.To = '%MAIL_FROM%';" ^
"$mail.Subject = '[긴급] CloudWatch disk_used_percent(/data) 지표 데이터 없음';" ^
"$mail.Body = '인스턴스 %INSTANCE_ID%의 disk_used_percent(/data) 지표에 수집된 데이터가 없습니다.';" ^
"$mail.Send()"
exit /b 1