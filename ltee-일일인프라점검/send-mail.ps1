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

$signHtml = if (Test-Path $MailSign) { Get-Content $MailSign -Raw } else { '' }

# $body = "<div>안녕하세요, 이노그리드 박지수입니다.</div>" +
#         "<div>$TitleDate 롯데에너지머티리얼즈 정보시스템 일일 점검서 전달 드립니다.</div>" +
#         "<div>첨부 파일 확인 부탁드립니다.</div><div>감사합니다.</div><div>박지수 드림.</div>" +
#         $signHtml

$body = @"
<html>
<head><meta charset="UTF-8"></head>
<body>
<div>안녕하세요, 이노그리드 박지수입니다.</div>
<div>$TitleDate 롯데에너지머티리얼즈 정보시스템 일일 점검서 전달 드립니다.</div>
<div>첨부 파일 확인 부탁드립니다.</div>
<div>감사합니다.</div>
<div>박지수 드림.</div>
$signHtml
</body>
</html>
"@

$ol   = New-Object -ComObject Outlook.Application
$mail = $ol.CreateItem(0)
$mail.To      = "jisoo449@naver.com"
# $mail.To      = "yangjeonghwan@lotte.net"
# $mail.CC      = "sunhwa_kim1@lotte.net;kjkang1@lotte.net;csp@innogrid.com;mjh@innogrid.com"
$mail.Subject = "[롯데에너지머티리얼즈] $TitleDate 정보시스템 일일 점검서 전달드립니다"
$mail.HTMLBody = $body
$mail.Attachments.Add($OutFile) | Out-Null
$mail.Send()