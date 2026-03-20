$file = "c:\Users\karl.lawrence\OneDrive - GCI Network Solutions\Desktop\Desktop Stuff\2 - Scripts & Automation\Scripts\AIWORK\DelprofPS\delprofPS.ps1"
$errors = $null
$tokens = [Management.Automation.Language.Parser]::Tokenize((Get-Content $file -Raw), [ref]$errors)
if ($errors) {
    foreach ($e in $errors) {
        if ($e.Extent.StartLineNumber -ge 840 -and $e.Extent.StartLineNumber -le 860) {
            Write-Output "Line $($e.Extent.StartLineNumber) Col $($e.Extent.StartColumnNumber): $($e.Message)"
        }
    }
}
$targetToken = $tokens | Where-Object { $_.Extent.StartLineNumber -eq 848 }
foreach ($t in $targetToken) {
    Write-Output "Token: $($t.Kind) Text: '$($t.Text)' Col: $($t.Extent.StartColumnNumber)"
}
