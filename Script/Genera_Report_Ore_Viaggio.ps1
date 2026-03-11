[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$InputPath,

    [Parameter(Position = 1)]
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Normalize-Header {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    return (($Value.Trim().ToLowerInvariant()) -replace '[^a-z0-9]', '')
}

function Normalize-Token {
    param([object]$Value)

    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToUpperInvariant()
}

function Escape-Xml {
    param([string]$Value)

    if ($null -eq $Value) { return '' }
    return [System.Security.SecurityElement]::Escape($Value)
}

function Write-Utf8File {
    param(
        [string]$Path,
        [string]$Content
    )

    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $encoding)
}

function Get-ExcelColumnName {
    param([int]$Index)

    $name = ''
    $n = $Index

    while ($n -gt 0) {
        $mod = ($n - 1) % 26
        $name = [char](65 + $mod) + $name
        $n = [int](($n - 1) / 26)
    }

    return $name
}

function Get-ColumnIndexFromCellReference {
    param([string]$CellReference)

    if ([string]::IsNullOrWhiteSpace($CellReference)) { return $null }

    $letters = ($CellReference -replace '[^A-Za-z]', '').ToUpperInvariant()
    if ($letters -eq '') { return $null }

    $index = 0
    foreach ($ch in $letters.ToCharArray()) {
        $index = ($index * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }

    return $index
}

function Get-ColumnName {
    param(
        [string[]]$Headers,
        [string[]]$Candidates,
        [string]$Label
    )

    foreach ($candidate in $Candidates) {
        foreach ($header in $Headers) {
            if ($header -ieq $candidate) { return $header }
        }
    }

    $normalizedHeaders = @{}
    foreach ($header in $Headers) {
        $normalizedHeaders[(Normalize-Header $header)] = $header
    }

    foreach ($candidate in $Candidates) {
        $key = Normalize-Header $candidate
        if ($normalizedHeaders.ContainsKey($key)) {
            return $normalizedHeaders[$key]
        }
    }

    throw "Colonna non trovata per '$Label'. Colonne disponibili: $($Headers -join ', ')"
}

function Find-FirstExistingColumn {
    param(
        [string[]]$Headers,
        [string[]]$Candidates
    )

    foreach ($candidate in $Candidates) {
        foreach ($header in $Headers) {
            if ($header -ieq $candidate) { return $header }
        }
    }

    $normalizedHeaders = @{}
    foreach ($header in $Headers) {
        $normalizedHeaders[(Normalize-Header $header)] = $header
    }

    foreach ($candidate in $Candidates) {
        $key = Normalize-Header $candidate
        if ($normalizedHeaders.ContainsKey($key)) {
            return $normalizedHeaders[$key]
        }
    }

    return $null
}

function Convert-ToDate {
    param([object]$Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [datetime]) {
        return ([datetime]$Value).Date
    }

    if ($Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal]) {
        try {
            return ([datetime]::FromOADate([double]$Value)).Date
        }
        catch {
            return $null
        }
    }

    $text = ([string]$Value).Trim()
    if ($text -eq '') { return $null }

    # Nei file xlsx la data puo arrivare come numero seriale Excel in formato stringa (es. 45931)
    $numericCandidate = $text -replace '\s', ''
    if ($numericCandidate -match '^-?\d+([.,]\d+)?$') {
        $asInvariant = $numericCandidate -replace ',', '.'
        $serial = 0.0
        if ([double]::TryParse($asInvariant, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$serial)) {
            try {
                return ([datetime]::FromOADate($serial)).Date
            }
            catch {
                # Se non e un seriale valido si continua con i parser testuali.
            }
        }
    }

    $formats = @(
        'dd/MM/yyyy',
        'd/M/yyyy',
        'dd-MM-yyyy',
        'd-M-yyyy',
        'yyyy-MM-dd',
        'dd/MM/yyyy HH:mm:ss',
        'd/M/yyyy H:mm',
        'dd-MM-yyyy HH:mm:ss',
        'd-M-yyyy H:mm'
    )

    $itCulture = [System.Globalization.CultureInfo]::GetCultureInfo('it-IT')
    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $styles = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces

    $parsed = [datetime]::MinValue

    if ([datetime]::TryParseExact($text, $formats, $itCulture, $styles, [ref]$parsed)) {
        return $parsed.Date
    }
    if ([datetime]::TryParse($text, $itCulture, $styles, [ref]$parsed)) {
        return $parsed.Date
    }
    if ([datetime]::TryParse($text, $invariant, $styles, [ref]$parsed)) {
        return $parsed.Date
    }

    return $null
}

function Convert-ToNumber {
    param([object]$Value)

    if ($null -eq $Value) { return 0.0 }

    if ($Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal]) {
        return [double]$Value
    }

    $text = ([string]$Value).Trim()
    if ($text -eq '') { return 0.0 }

    $candidate = $text -replace '\s', ''

    if ($candidate -match '^-?\d{1,3}(\.\d{3})*(,\d+)?$') {
        $candidate = $candidate -replace '\.', ''
        $candidate = $candidate -replace ',', '.'
    }
    elseif ($candidate -match '^-?\d+,\d+$') {
        $candidate = $candidate -replace ',', '.'
    }

    $number = 0.0
    if ([double]::TryParse($candidate, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::GetCultureInfo('it-IT'), [ref]$number)) {
        return $number
    }

    return 0.0
}

function Import-CsvAuto {
    param([string]$Path)

    $firstLine = Get-Content -Path $Path -TotalCount 1
    $delimiter = ','
    if ($firstLine -match ';') { $delimiter = ';' }

    return Import-Csv -Path $Path -Delimiter $delimiter
}

function Read-ZipEntryText {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryPath,
        [switch]$AllowMissing
    )

    $normalized = $EntryPath -replace '\\', '/'
    $alternate = $EntryPath -replace '/', '\'

    $entry = $Archive.Entries |
        Where-Object {
            $_.FullName -ieq $EntryPath -or
            $_.FullName -ieq $normalized -or
            $_.FullName -ieq $alternate
        } |
        Select-Object -First 1
    if ($null -eq $entry) {
        if ($AllowMissing) { return $null }
        throw "Voce ZIP non trovata: $EntryPath"
    }

    $stream = $null
    $reader = $null
    try {
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        return $reader.ReadToEnd()
    }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-XlsxCellValue {
    param(
        [System.Xml.XmlNode]$CellNode,
        [string[]]$SharedStrings
    )

    $cellType = if ($CellNode.Attributes['t']) { $CellNode.Attributes['t'].Value } else { '' }

    if ($cellType -eq 'inlineStr') {
        $textNodes = $CellNode.SelectNodes(".//*[local-name()='t']")
        if ($null -eq $textNodes -or $textNodes.Count -eq 0) { return '' }
        $parts = @()
        foreach ($node in $textNodes) { $parts += $node.InnerText }
        return ($parts -join '')
    }

    $vNode = $CellNode.SelectSingleNode("./*[local-name()='v']")
    $raw = if ($null -eq $vNode) { '' } else { $vNode.InnerText }

    if ($cellType -eq 's') {
        $index = 0
        if ([int]::TryParse($raw, [ref]$index) -and $index -ge 0 -and $index -lt $SharedStrings.Count) {
            return $SharedStrings[$index]
        }
        return ''
    }

    return $raw
}

function Import-XlsxOpenXml {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)

        $workbookText = Read-ZipEntryText -Archive $archive -EntryPath 'xl/workbook.xml'
        [xml]$workbookXml = $workbookText

        $sheetNode = $workbookXml.SelectSingleNode("//*[local-name()='sheets']/*[local-name()='sheet'][1]")
        if ($null -eq $sheetNode) {
            throw 'Foglio Excel non trovato nel file input.'
        }

        $sheetRelId = $null
        if ($sheetNode.Attributes['r:id']) {
            $sheetRelId = [string]$sheetNode.Attributes['r:id'].Value
        }
        if ([string]::IsNullOrWhiteSpace($sheetRelId)) {
            $sheetElement = [System.Xml.XmlElement]$sheetNode
            $sheetRelId = $sheetElement.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        }
        if ([string]::IsNullOrWhiteSpace($sheetRelId)) {
            throw 'Relazione del primo foglio non trovata.'
        }

        $relsText = Read-ZipEntryText -Archive $archive -EntryPath 'xl/_rels/workbook.xml.rels'
        [xml]$relsXml = $relsText
        $relationshipNode = $relsXml.SelectNodes("//*[local-name()='Relationship']") |
            Where-Object { $_.Id -eq $sheetRelId } |
            Select-Object -First 1

        if ($null -eq $relationshipNode) {
            throw 'Target del primo foglio non trovato nelle relazioni.'
        }

        $target = [string]$relationshipNode.Target
        $baseUri = [System.Uri]::new('http://dummy/xl/workbook.xml')
        $sheetPath = [System.Uri]::new($baseUri, $target).AbsolutePath.TrimStart('/')

        $sheetText = Read-ZipEntryText -Archive $archive -EntryPath $sheetPath
        [xml]$sheetXml = $sheetText

        $sharedStrings = @()
        $sharedText = Read-ZipEntryText -Archive $archive -EntryPath 'xl/sharedStrings.xml' -AllowMissing
        if ($null -ne $sharedText) {
            [xml]$sharedXml = $sharedText
            $siNodes = $sharedXml.SelectNodes("//*[local-name()='si']")
            foreach ($si in $siNodes) {
                $textNodes = $si.SelectNodes(".//*[local-name()='t']")
                if ($null -eq $textNodes -or $textNodes.Count -eq 0) {
                    $sharedStrings += ''
                }
                else {
                    $parts = @()
                    foreach ($txt in $textNodes) { $parts += $txt.InnerText }
                    $sharedStrings += ($parts -join '')
                }
            }
        }

        $rowNodes = $sheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")
        if ($null -eq $rowNodes -or $rowNodes.Count -eq 0) {
            return @()
        }

        $headerRow = $rowNodes[0]
        $headerCells = $headerRow.SelectNodes("./*[local-name()='c']")
        if ($null -eq $headerCells -or $headerCells.Count -eq 0) {
            return @()
        }

        $headerMap = @{}
        $seqIndex = 1
        foreach ($cell in $headerCells) {
            $ref = if ($cell.Attributes['r']) { [string]$cell.Attributes['r'].Value } else { '' }
            $colIndex = Get-ColumnIndexFromCellReference $ref
            if ($null -eq $colIndex) { $colIndex = $seqIndex }

            $headerValue = ([string](Get-XlsxCellValue -CellNode $cell -SharedStrings $sharedStrings)).Trim()
            if ($headerValue -eq '') {
                $headerValue = "Colonna_$colIndex"
            }

            $headerMap[[string]$colIndex] = $headerValue
            $seqIndex = [Math]::Max($seqIndex + 1, $colIndex + 1)
        }

        $rows = New-Object System.Collections.Generic.List[object]

        for ($r = 1; $r -lt $rowNodes.Count; $r++) {
            $rowNode = $rowNodes[$r]
            $cells = $rowNode.SelectNodes("./*[local-name()='c']")

            $valueByColumn = @{}
            $seqCol = 1
            foreach ($cell in $cells) {
                $ref = if ($cell.Attributes['r']) { [string]$cell.Attributes['r'].Value } else { '' }
                $colIndex = Get-ColumnIndexFromCellReference $ref
                if ($null -eq $colIndex) { $colIndex = $seqCol }

                $valueByColumn[[string]$colIndex] = Get-XlsxCellValue -CellNode $cell -SharedStrings $sharedStrings
                $seqCol = [Math]::Max($seqCol + 1, $colIndex + 1)
            }

            $obj = [ordered]@{}
            $isEmpty = $true
            foreach ($colKey in ($headerMap.Keys | Sort-Object { [int]$_ })) {
                $header = $headerMap[$colKey]
                $value = if ($valueByColumn.ContainsKey($colKey)) { $valueByColumn[$colKey] } else { $null }
                if ($null -ne $value -and ([string]$value).Trim() -ne '') {
                    $isEmpty = $false
                }
                $obj[$header] = $value
            }

            if (-not $isEmpty) {
                $rows.Add([pscustomobject]$obj)
            }
        }

        return $rows.ToArray()
    }
    finally {
        if ($null -ne $archive) { $archive.Dispose() }
    }
}

function Import-ExcelWithCom {
    param([string]$Path)

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $usedRange = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($Path, 0, $true)
        $worksheet = $workbook.Worksheets.Item(1)
        $usedRange = $worksheet.UsedRange

        $rowCount = [int]$usedRange.Rows.Count
        $colCount = [int]$usedRange.Columns.Count

        if ($rowCount -lt 2 -or $colCount -lt 1) {
            return @()
        }

        $headers = @()
        for ($c = 1; $c -le $colCount; $c++) {
            $header = [string]$worksheet.Cells.Item(1, $c).Text
            $header = $header.Trim()
            if ($header -eq '') {
                $header = "Colonna_$c"
            }
            $headers += $header
        }

        $rows = New-Object System.Collections.Generic.List[object]
        for ($r = 2; $r -le $rowCount; $r++) {
            $obj = [ordered]@{}
            $isEmpty = $true

            for ($c = 1; $c -le $colCount; $c++) {
                $header = $headers[$c - 1]
                $value = $worksheet.Cells.Item($r, $c).Value2
                if ($null -ne $value -and ([string]$value).Trim() -ne '') {
                    $isEmpty = $false
                }
                $obj[$header] = $value
            }

            if (-not $isEmpty) {
                $rows.Add([pscustomobject]$obj)
            }
        }

        return $rows.ToArray()
    }
    finally {
        if ($null -ne $usedRange) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) }
        if ($null -ne $worksheet) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) }
        if ($null -ne $workbook) {
            $workbook.Close($false)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }
        if ($null -ne $excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Import-TabularData {
    param([string]$Path)

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()

    switch ($extension) {
        '.csv' {
            return @(Import-CsvAuto -Path $Path)
        }
        '.xlsx' {
            return @(Import-XlsxOpenXml -Path $Path)
        }
        '.xls' {
            throw 'Formato .xls non supportato in modalita portabile. Converti prima il file in .xlsx o .csv.'
        }
        default {
            throw "Formato non supportato: $extension"
        }
    }
}

function Export-XlsxOpenXml {
    param(
        [object[]]$Rows,
        [string]$Path
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("report_ore_viaggio_" + [Guid]::NewGuid().ToString('N'))

    try {
        New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null
        New-Item -Path (Join-Path $tempRoot '_rels') -ItemType Directory -Force | Out-Null
        New-Item -Path (Join-Path $tempRoot 'xl') -ItemType Directory -Force | Out-Null
        New-Item -Path (Join-Path $tempRoot 'xl\_rels') -ItemType Directory -Force | Out-Null
        New-Item -Path (Join-Path $tempRoot 'xl\worksheets') -ItemType Directory -Force | Out-Null

        $contentTypesLines = @(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
            '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
            '  <Default Extension="xml" ContentType="application/xml"/>',
            '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
            '</Types>'
        )
        $rootRelsLines = @(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
            '</Relationships>'
        )
        $workbookLines = @(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            '  <sheets>',
            '    <sheet name="Report" sheetId="1" r:id="rId1"/>',
            '  </sheets>',
            '</workbook>'
        )
        $workbookRelsLines = @(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
            '</Relationships>'
        )

        Write-Utf8File -Path (Join-Path $tempRoot '[Content_Types].xml') -Content ($contentTypesLines -join "`r`n")
        Write-Utf8File -Path (Join-Path $tempRoot '_rels\.rels') -Content ($rootRelsLines -join "`r`n")
        Write-Utf8File -Path (Join-Path $tempRoot 'xl\workbook.xml') -Content ($workbookLines -join "`r`n")
        Write-Utf8File -Path (Join-Path $tempRoot 'xl\_rels\workbook.xml.rels') -Content ($workbookRelsLines -join "`r`n")

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        [void]$sb.AppendLine('  <sheetData>')

        if (-not $Rows -or $Rows.Count -eq 0) {
            throw 'Nessun dato da esportare nel file di output.'
        }

        $headers = @($Rows[0].PSObject.Properties.Name)
        $numericColumns = @('somma_totale', 'Ore_Viaggio')

        [void]$sb.AppendLine('    <row r="1">')
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $colName = Get-ExcelColumnName ($c + 1)
            $cellRef = "$colName`1"
            $headerEscaped = Escape-Xml $headers[$c]
            [void]$sb.AppendLine(('      <c r="{0}" t="inlineStr"><is><t>{1}</t></is></c>' -f $cellRef, $headerEscaped))
        }
        [void]$sb.AppendLine('    </row>')

        $rowIndex = 2
        $invariant = [System.Globalization.CultureInfo]::InvariantCulture

        foreach ($row in $Rows) {
            [void]$sb.AppendLine(('    <row r="{0}">' -f $rowIndex))

            for ($c = 0; $c -lt $headers.Count; $c++) {
                $header = $headers[$c]
                $colName = Get-ExcelColumnName ($c + 1)
                $cellRef = "$colName$rowIndex"
                $value = $row.PSObject.Properties[$header].Value

                if ($numericColumns -contains $header) {
                    $numText = (Convert-ToNumber $value).ToString($invariant)
                    [void]$sb.AppendLine(('      <c r="{0}"><v>{1}</v></c>' -f $cellRef, $numText))
                    continue
                }

                if ($value -is [datetime]) {
                    $textValue = ([datetime]$value).ToString('dd/MM/yyyy')
                }
                else {
                    $textValue = [string]$value
                }

                $textEscaped = Escape-Xml $textValue
                [void]$sb.AppendLine(('      <c r="{0}" t="inlineStr"><is><t>{1}</t></is></c>' -f $cellRef, $textEscaped))
            }

            [void]$sb.AppendLine('    </row>')

            $rowIndex++
        }

        [void]$sb.AppendLine('  </sheetData>')
        [void]$sb.AppendLine('</worksheet>')

        Write-Utf8File -Path (Join-Path $tempRoot 'xl\worksheets\sheet1.xml') -Content $sb.ToString()

        $outputDir = Split-Path -Parent $Path
        if (-not [string]::IsNullOrWhiteSpace($outputDir) -and -not (Test-Path $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        }

        if (Test-Path $Path) {
            Remove-Item -Path $Path -Force
        }

        $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            $files = Get-ChildItem -Path $tempRoot -Recurse -File
            foreach ($file in $files) {
                $relative = $file.FullName.Substring($tempRoot.Length).TrimStart('\', '/')
                $entryName = $relative -replace '\\', '/'
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                    $archive,
                    $file.FullName,
                    $entryName,
                    [System.IO.Compression.CompressionLevel]::Optimal
                ) | Out-Null
            }
        }
        finally {
            $archive.Dispose()
        }
    }
    finally {
        if (Test-Path $tempRoot) {
            Remove-Item -Path $tempRoot -Recurse -Force
        }
    }
}

function Resolve-InputPath {
    param([string]$InputDirectory)

    $candidates = @(
        Get-ChildItem -Path $InputDirectory -File |
            Where-Object { $_.Extension -in @('.xlsx', '.xls', '.csv') -and -not $_.Name.StartsWith('~$') }
    )

    if (-not $candidates -or $candidates.Count -eq 0) {
        throw 'Nessun file di input trovato nella cartella Input. Metti li il report (.xlsx/.csv).'
    }

    $ranked = $candidates |
        ForEach-Object {
            $name = $_.Name.ToLowerInvariant()
            $score = 0

            if ($name -match 'rendicont') { $score += 5 }
            if ($name -match 'ottobre') { $score += 3 }
            if ($name -match '2025') { $score += 2 }
            if ($name -match 'ore') { $score += 1 }
            if ($_.Extension -eq '.xlsx') { $score += 1 }

            [pscustomobject]@{
                File = $_
                Score = $score
            }
        } |
        Sort-Object -Property @{ Expression = 'Score'; Descending = $true }, @{ Expression = { $_.File.LastWriteTimeUtc }; Descending = $true }

    return $ranked[0].File.FullName
}

$scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$rootDirectory = if ((Split-Path -Leaf $scriptDirectory).ToLowerInvariant() -eq 'script') {
    Split-Path -Parent $scriptDirectory
}
else {
    $scriptDirectory
}

$inputDirectory = Join-Path $rootDirectory 'Input'
$outputDirectory = Join-Path $rootDirectory 'Output'
New-Item -Path $inputDirectory -ItemType Directory -Force | Out-Null
New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null

if ([string]::IsNullOrWhiteSpace($InputPath)) {
    $InputPath = Resolve-InputPath -InputDirectory $inputDirectory
}
elseif (-not [System.IO.Path]::IsPathRooted($InputPath)) {
    $candidateFromInput = Join-Path $inputDirectory $InputPath
    $candidateFromRoot = Join-Path $rootDirectory $InputPath

    if (Test-Path $candidateFromInput) {
        $InputPath = $candidateFromInput
    }
    elseif (Test-Path $candidateFromRoot) {
        $InputPath = $candidateFromRoot
    }
    else {
        $InputPath = $candidateFromInput
    }
}

if (-not (Test-Path $InputPath)) {
    throw "File di input non trovato: $InputPath"
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $outputDirectory 'Report_Ore_Viaggio_Filtrato.xlsx'
}
elseif (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path $outputDirectory $OutputPath
}

if ([System.IO.Path]::GetExtension($OutputPath).ToLowerInvariant() -ne '.xlsx') {
    $OutputPath = [System.IO.Path]::ChangeExtension($OutputPath, '.xlsx')
}

Write-Host "Input rilevato: $InputPath"
Write-Host "Output previsto: $OutputPath"

$rows = @(Import-TabularData -Path $InputPath)
if (-not $rows -or $rows.Count -eq 0) {
    throw 'Nessun dato trovato nel file di input.'
}

$headers = @($rows[0].PSObject.Properties.Name)

$colRepartoDescr = Find-FirstExistingColumn -Headers $headers -Candidates @('Descr.Reparto', 'Descr Reparto', 'Descrizione Reparto')
$colRepartoCode = Find-FirstExistingColumn -Headers $headers -Candidates @('Reparto', 'Cod. Reparto', 'Codice Reparto')
if ($null -eq $colRepartoDescr -and $null -eq $colRepartoCode) {
    throw "Colonna reparto non trovata. Colonne disponibili: $($headers -join ', ')"
}
$colQta = Get-ColumnName -Headers $headers -Candidates @('Qta', 'Quantita', 'Quantità', 'Ore') -Label 'Qta'
$colCodArgomento = Get-ColumnName -Headers $headers -Candidates @('Cod. Argomento', 'Cod Argomento', 'Codice Argomento', 'Argomento') -Label 'Cod. Argomento'
$colProgetto = Get-ColumnName -Headers $headers -Candidates @('Progetto', 'Commessa') -Label 'Progetto'
$colData = Get-ColumnName -Headers $headers -Candidates @('Data', 'Giorno', 'Data Lavoro', 'Data Consuntivo', 'Data Registrazione') -Label 'Data'
$colCodiceAzienda = Get-ColumnName -Headers $headers -Candidates @('Codice Azienda', 'Codice azienda', 'Cod. Azienda') -Label 'Codice Azienda'
$colDescrAzienda = Get-ColumnName -Headers $headers -Candidates @('Descr. Azienda', 'Descr Azienda', 'Descrizione Azienda') -Label 'Descr. Azienda'
$colCodiceDipendente = Get-ColumnName -Headers $headers -Candidates @('Codice dipendente', 'Codice Dipendente', 'Cod. Dipendente', 'Matricola') -Label 'Codice dipendente'

$colManutentore = Find-FirstExistingColumn -Headers $headers -Candidates @('Manutentore', 'Descr.Risorsa', 'Descr Risorsa', 'Nominativo', 'Dipendente', 'Risorsa', 'Tecnico', 'Operatore')
if ($null -eq $colManutentore) {
    throw "Colonna manutentore non trovata. Colonne disponibili: $($headers -join ', ')"
}

$groups = @{}
$filteredRows = 0

foreach ($row in $rows) {
    $repartoDescr = if ($null -ne $colRepartoDescr) { Normalize-Token $row.$colRepartoDescr } else { '' }
    $repartoCode = if ($null -ne $colRepartoCode) { Normalize-Token $row.$colRepartoCode } else { '' }

    if ($repartoDescr -ne 'MANUTENTORI' -and $repartoCode -ne 'MAN') {
        continue
    }

    $dateValue = Convert-ToDate $row.$colData
    if ($null -eq $dateValue) {
        continue
    }

    $manutentore = ([string]$row.$colManutentore).Trim()
    if ([string]::IsNullOrWhiteSpace($manutentore)) {
        continue
    }

    $codiceAzienda = ([string]$row.$colCodiceAzienda).Trim()
    $descrAzienda = ([string]$row.$colDescrAzienda).Trim()
    $codiceDipendente = ([string]$row.$colCodiceDipendente).Trim()

    if ([string]::IsNullOrWhiteSpace($codiceDipendente)) {
        $codiceDipendente = $manutentore
    }

    $qta = Convert-ToNumber $row.$colQta
    $key = '{0}|{1}|{2}' -f $codiceDipendente.ToUpperInvariant(), $manutentore.ToUpperInvariant(), $dateValue.ToString('yyyy-MM-dd')

    if (-not $groups.ContainsKey($key)) {
        $groups[$key] = [ordered]@{
            'Codice Azienda' = $codiceAzienda
            'Descr. Azienda' = $descrAzienda
            'Codice dipendente' = $codiceDipendente
            Manutentore = $manutentore
            Data = $dateValue
            somma_totale = 0.0
            Ore_Viaggio = 0.0
        }
    }

    $groups[$key].somma_totale += $qta

    $codArg = Normalize-Token $row.$colCodArgomento
    $progetto = Normalize-Token $row.$colProgetto
    if ($codArg -eq 'CHIUSURA' -or $progetto -eq 'COMMESSA') {
        $groups[$key].Ore_Viaggio += $qta
    }

    $filteredRows++
}

if ($groups.Count -eq 0) {
    throw 'Nessuna riga valida trovata dopo il filtro reparto (Descr.Reparto = MANUTENTORI o Reparto = MAN).'
}

$result = @(
    $groups.Values |
        ForEach-Object {
            [pscustomobject]@{
                'Codice Azienda' = $_.'Codice Azienda'
                'Descr. Azienda' = $_.'Descr. Azienda'
                'Codice dipendente' = $_.'Codice dipendente'
                Manutentore = $_.Manutentore
                Data = $_.Data
                somma_totale = [math]::Round($_.somma_totale, 2)
                Ore_Viaggio = [math]::Round($_.Ore_Viaggio, 2)
            }
        } |
        Sort-Object -Property 'Codice dipendente', Manutentore, Data
)

Export-XlsxOpenXml -Rows $result -Path $OutputPath

Write-Host "Righe MANUTENTORI elaborate: $filteredRows"
Write-Host "Righe aggregate generate: $($result.Count)"
Write-Host "Report creato: $OutputPath"
