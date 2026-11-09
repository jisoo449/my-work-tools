param(
  [string]$MailSign,
  [string]$TitleDate,
  [string]$OutFile,
  [string]$MailFrom
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Write-Host "MailSign: $MailSign"
# Write-Host "TitleDate: $TitleDate"
# Write-Host "OutFile: $OutFile"
# Write-Host "MailFrom: $MailFrom"
# pause

# 수신자/발신자 이름은 같은 폴더의 config.env 에서 읽음 (git 제외)
$cfg = @{}
Get-Content (Join-Path $PSScriptRoot 'config.env') -Encoding UTF8 | Where-Object { $_ -match '^\s*[^#].*=' } | ForEach-Object {
  $k, $v = $_ -split '=', 2
  $cfg[$k.Trim()] = $v.Trim()
}

$signHtml = if (Test-Path $MailSign) { Get-Content $MailSign -Raw } else { '' }

# $body = "<div>안녕하세요, $($cfg.COMPANY_NAME) $($cfg.SENDER_NAME)입니다.</div>" +
#         "<div>$TitleDate 롯데에너지머티리얼즈 정보시스템 일일 점검서 전달 드립니다.</div>" +
#         "<div>첨부 파일 확인 부탁드립니다.</div><div>감사합니다.</div><div>$($cfg.SENDER_NAME) 드림.</div>" +
#         $signHtml

$body = @"
<html>
<head><meta charset="UTF-8"></head>
<body>
<div>안녕하세요, $($cfg.COMPANY_NAME) $($cfg.SENDER_NAME)입니다.</div>
<div>$TitleDate 롯데에너지머티리얼즈 정보시스템 일일 점검서 전달 드립니다.</div>
<div>첨부 파일 확인 부탁드립니다.</div>
<div>감사합니다.</div>
<div>$($cfg.SENDER_NAME) 드림.</div>
$signHtml
</body>
</html>
"@

$ol   = New-Object -ComObject Outlook.Application
$mail = $ol.CreateItem(0)
$mail.To      = $cfg.TEST_MAIL_TO
# $mail.To      = $cfg.MAIL_TO
# $mail.CC      = $cfg.MAIL_CC -replace ',', ';'
$mail.Subject = "[롯데에너지머티리얼즈] $TitleDate 정보시스템 일일 점검서 전달드립니다"
$mail.HTMLBody = $body
$mail.Attachments.Add($OutFile) | Out-Null
$mail.Send()