param(
    [string]$WorkbookPath,
    [string]$SourcePath
)

$ErrorActionPreference = "Stop"

function Get-ExcelApplication {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        return $excel
    }
}

function Get-WorkbookByPath($excel, $targetPath) {
    $targetFull = [IO.Path]::GetFullPath($targetPath)
    foreach ($wb in $excel.Workbooks) {
        try {
            if ([string]::Equals(
                [IO.Path]::GetFullPath($wb.FullName),
                $targetFull,
                [StringComparison]::OrdinalIgnoreCase
            )) {
                return $wb
            }
        } catch {}
    }
    return $null
}

function Get-ComponentByName($vbProject, $name) {
    foreach ($component in $vbProject.VBComponents) {
        if ([string]::Equals(
            $component.Name,
            $name,
            [StringComparison]::OrdinalIgnoreCase
        )) {
            return $component
        }
    }
    return $null
}

function Remove-ExistingComponent($vbProject, $name) {
    $existing = Get-ComponentByName $vbProject $name
    if ($null -eq $existing) { return }
    if ([int]$existing.Type -eq 100) {
        throw "Cannot replace document module '$name'."
    }
    $vbProject.VBComponents.Remove($existing)
}

function Read-Utf8Text($filePath) {
    $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
    return [IO.File]::ReadAllText($filePath, $utf8)
}

function Strip-VbaExportHeader($text) {
    $text = [Regex]::Replace(
        $text,
        '(?ms)^\s*VERSION[^\r\n]*\r?\nBEGIN.*?^\s*END\s*\r?\n',
        ''
    )
    $text = [Regex]::Replace(
        $text,
        '(?m)^\s*Attribute\s+[^\r\n]*\r?\n?',
        ''
    )
    return $text.TrimStart()
}

function Convert-SourceToExcelAnsi($sourcePath, $tempPath) {
    $text = Read-Utf8Text $sourcePath
    $ansi = [Text.Encoding]::Default
    [IO.File]::WriteAllBytes($tempPath, $ansi.GetBytes($text))
}

try {
    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
        throw "Workbook does not exist: $WorkbookPath"
    }
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        throw "Source file does not exist: $SourcePath"
    }

    $extension = [IO.Path]::GetExtension($SourcePath).ToLowerInvariant()
    $componentName = [IO.Path]::GetFileNameWithoutExtension($SourcePath)

    switch ($extension) {
        ".bas" { $expectedType = 1 }
        ".cls" { $expectedType = 2 }
        ".frm" { $expectedType = 3 }
        default { throw "Unsupported VBA extension: $extension" }
    }

    $excel = Get-ExcelApplication
    $excel.Visible = $true
    $null = $excel.VBE

    $workbook = Get-WorkbookByPath $excel $WorkbookPath
    if ($null -eq $workbook) {
        $workbook = $excel.Workbooks.Open($WorkbookPath)
    }
    if ($workbook.ReadOnly) {
        throw "Workbook is read-only: $WorkbookPath"
    }

    $vbProject = $workbook.VBProject
    $tempImport = $null

    try {
        if ($extension -eq ".frm") {
            $tempImport = Join-Path $env:TEMP (
                "vba-cli-" + [Guid]::NewGuid().ToString() + ".frm"
            )
            Convert-SourceToExcelAnsi $SourcePath $tempImport
            Remove-ExistingComponent $vbProject $componentName
            $component = $vbProject.VBComponents.Import($tempImport)
        }
        else {
            Remove-ExistingComponent $vbProject $componentName
            $component = $vbProject.VBComponents.Add($expectedType)
            $component.Name = $componentName
            $code = Strip-VbaExportHeader (Read-Utf8Text $SourcePath)
            if (-not [string]::IsNullOrWhiteSpace($code)) {
                $component.CodeModule.AddFromString($code)
            }
        }

        if ($null -eq $component) {
            throw "Could not create/import component '$componentName'."
        }

        if ([int]$component.Type -ne $expectedType) {
            $actualType = [int]$component.Type
            $vbProject.VBComponents.Remove($component)
            throw "Component '$componentName' has type $actualType, expected $expectedType."
        }

        $workbook.Save()
        Write-Output ("OK|" + $componentName + "|" + $expectedType)
    }
    finally {
        if ($null -ne $tempImport) {
            Remove-Item -LiteralPath $tempImport -Force -ErrorAction SilentlyContinue
        }
    }
}
catch {
    [Console]::Error.WriteLine(("VBA_IMPORT_ERROR: " + $_.Exception.Message))
    exit 1
}
