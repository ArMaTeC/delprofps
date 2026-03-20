$text = Get-Content "c:\Users\karl.lawrence\OneDrive - GCI Network Solutions\Desktop\Desktop Stuff\2 - Scripts & Automation\Scripts\AIWORK\DelprofPS\delprofPS.ps1" -Raw
$open = 0
$close = 0
for ($i = 0; $i -lt $text.Length; $i++) {
    if ($text[$i] -eq '{') { $open++ }
    if ($text[$i] -eq '}') { $close++ }
}
Write-Output "Open: $open, Close: $close"
if ($open -ne $close) {
    Write-Output "Mismatch found!"
} else {
    Write-Output "Matches."
}
