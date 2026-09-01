param(
    [Parameter(Mandatory=$true)]
    [string]$ProjectRoot
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$ExcelProcesses = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($ExcelProcesses.Count -gt 0) {
    throw "Close every open Excel window and run install.cmd again."
}

$SecurityKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$PropertyName = "AccessVBOM"
$KeyExisted = Test-Path $SecurityKey
$ValueExisted = $false
$OldValue = $null

if ($KeyExisted) {
    try {
        $OldValue = (Get-ItemProperty -Path $SecurityKey -Name $PropertyName -ErrorAction Stop).$PropertyName
        $ValueExisted = $true
    }
    catch {
        $ValueExisted = $false
    }
}
else {
    New-Item -Path $SecurityKey -Force | Out-Null
}

try {
    Set-ItemProperty -Path $SecurityKey -Name $PropertyName -Type DWord -Value 1
    $Builder = Join-Path $ProjectRoot "scripts\build_excel_template.vbs"
    & cscript.exe //nologo $Builder $ProjectRoot
    if ($LASTEXITCODE -ne 0) {
        throw "Excel template builder returned code $LASTEXITCODE."
    }

    $Output = Join-Path $ProjectRoot "MediaCatalog_template.xlsm"
    if (-not (Test-Path $Output)) {
        throw "Excel did not create MediaCatalog_template.xlsm."
    }
    Write-Host "Created $Output"
}
finally {
    if ($ValueExisted) {
        Set-ItemProperty -Path $SecurityKey -Name $PropertyName -Type DWord -Value $OldValue
    }
    else {
        Remove-ItemProperty -Path $SecurityKey -Name $PropertyName -ErrorAction SilentlyContinue
        if (-not $KeyExisted) {
            $Remaining = @(Get-ItemProperty -Path $SecurityKey -ErrorAction SilentlyContinue).PSObject.Properties |
                Where-Object { $_.Name -notmatch "^PS" }
            if ($Remaining.Count -eq 0) {
                Remove-Item -Path $SecurityKey -ErrorAction SilentlyContinue
            }
        }
    }
}
