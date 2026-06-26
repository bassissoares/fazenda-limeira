param(
    [string]$Source = "IATF TETF 25-26.xlsx",
    [string]$OutputDir = "data"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Read-ZipEntryText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$Name
    )

    $entry = $Zip.GetEntry($Name)
    if ($null -eq $entry) {
        return $null
    }

    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
        return $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }
}

function Get-ColIndex {
    param([string]$CellRef)

    $letters = ([regex]::Match($CellRef, "^[A-Z]+")).Value
    $value = 0
    foreach ($char in $letters.ToCharArray()) {
        $value = ($value * 26) + ([int][char]$char - [int][char]"A" + 1)
    }
    return $value
}

function Get-ChildText {
    param(
        [System.Xml.XmlElement]$Element,
        [string]$LocalName
    )

    foreach ($node in $Element.ChildNodes) {
        if ($node.LocalName -eq $LocalName) {
            return [string]$node.InnerText
        }
    }

    return ""
}

function Get-CellValue {
    param(
        [System.Xml.XmlElement]$Cell,
        [string[]]$SharedStrings
    )

    $cellType = $Cell.GetAttribute("t")

    if ($cellType -eq "s") {
        $value = Get-ChildText -Element $Cell -LocalName "v"
        if ([string]::IsNullOrWhiteSpace($value)) {
            return ""
        }
        return ($SharedStrings[[int]$value]).Trim()
    }

    if ($cellType -eq "inlineStr") {
        $texts = @()
        foreach ($node in $Cell.is.ChildNodes) {
            if ($node.LocalName -eq "t") {
                $texts += $node.InnerText
            }
        }
        return (($texts -join "")).Trim()
    }

    $rawValue = Get-ChildText -Element $Cell -LocalName "v"
    if (-not [string]::IsNullOrWhiteSpace($rawValue)) {
        return $rawValue.Trim()
    }

    return ""
}

function Find-HeaderIndex {
    param(
        [string[]]$Headers,
        [string[]]$Patterns
    )

    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $header = $Headers[$i].ToUpperInvariant()
        foreach ($pattern in $Patterns) {
            if ($header -like $pattern) {
                return $i
            }
        }
    }

    return -1
}

function Normalize-Name {
    param([string]$Value)

    $text = ($Value -replace "\s+", " ").Trim().ToUpperInvariant()
    $map = @{
        "HANKOK"     = "HANCOCK"
        "HANKOCK"    = "HANCOCK"
        "BRUTÃO ARF" = "BRUTAO ARF"
        "XARP "      = "XARP"
        "TATICO"     = "TATICO SN"
        "FENOL S"    = "FENOL"
        "ARGOS"      = "ARGOS SN"
        "XARP COL"   = "XARP COL"
    }

    if ($map.ContainsKey($text)) {
        return $map[$text]
    }

    return $text
}

function Normalize-Dg {
    param([string]$Value)

    $text = ($Value -replace "\s+", " ").Trim().ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ""
    }

    if ($text -match "^P\+") {
        return "P+"
    }

    if ($text -eq "X") {
        return "X"
    }

    if ($text -match "^V") {
        return "V"
    }

    return $text
}

function Get-Rate {
    param([int]$Count, [int]$Total)

    if ($Total -le 0) {
        return 0
    }

    return [math]::Round(($Count * 100.0) / $Total, 1)
}

function Convert-ToDateText {
    param([string]$Header)

    $match = [regex]::Match($Header, "(\d{1,2}/\d{1,2}(?:/\d{2,4})?)")
    if (-not $match.Success) {
        return ""
    }

    return $match.Groups[1].Value
}

$sourcePath = Join-Path (Get-Location) $Source
if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "Arquivo fonte nao encontrado: $sourcePath"
}

$outputPath = Join-Path (Get-Location) $OutputDir
New-Item -ItemType Directory -Force -Path $outputPath | Out-Null

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($sourcePath)

try {
    $sharedStrings = @()
    $sharedText = Read-ZipEntryText -Zip $zip -Name "xl/sharedStrings.xml"
    if ($sharedText) {
        [xml]$sharedXml = $sharedText
        $sharedNs = [System.Xml.XmlNamespaceManager]::new($sharedXml.NameTable)
        $sharedNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        foreach ($si in $sharedXml.SelectNodes("//x:si", $sharedNs)) {
            $sharedStrings += (($si.SelectNodes(".//x:t", $sharedNs) | ForEach-Object { $_.InnerText }) -join "")
        }
    }

    [xml]$workbook = Read-ZipEntryText -Zip $zip -Name "xl/workbook.xml"
    [xml]$rels = Read-ZipEntryText -Zip $zip -Name "xl/_rels/workbook.xml.rels"
    $wbNs = [System.Xml.XmlNamespaceManager]::new($workbook.NameTable)
    $wbNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $wbNs.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    $relMap = @{}
    foreach ($rel in $rels.Relationships.Relationship) {
        $relMap[$rel.Id] = $rel.Target
    }

    $records = @()
    $sheetMetas = @()

    foreach ($sheet in $workbook.SelectNodes("//x:sheet", $wbNs)) {
        $rid = $sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        $target = $relMap[$rid]
        if ($target -notlike "worksheets/*") {
            $target = "worksheets/$target"
        }

        [xml]$worksheet = Read-ZipEntryText -Zip $zip -Name "xl/$target"
        $wsNs = [System.Xml.XmlNamespaceManager]::new($worksheet.NameTable)
        $wsNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

        $matrix = @()
        foreach ($row in $worksheet.SelectNodes("//x:sheetData/x:row", $wsNs)) {
            $cells = @{}
            $maxColumn = 0
            foreach ($cell in $row.SelectNodes("x:c", $wsNs)) {
                $columnIndex = Get-ColIndex -CellRef $cell.r
                if ($columnIndex -gt $maxColumn) {
                    $maxColumn = $columnIndex
                }
                $cells[$columnIndex] = Get-CellValue -Cell $cell -SharedStrings $sharedStrings
            }

            $rowValues = @()
            for ($i = 1; $i -le $maxColumn; $i++) {
                if ($cells.ContainsKey($i)) {
                    $rowValues += $cells[$i]
                }
                else {
                    $rowValues += ""
                }
            }
            $matrix += ,$rowValues
        }

        if ($matrix.Count -eq 0) {
            continue
        }

        [string[]]$headers = $matrix[0]
        $dgIndex = Find-HeaderIndex -Headers $headers -Patterns @("*DG*")
        $bezerroIndex = Find-HeaderIndex -Headers $headers -Patterns @("*BEZERRO*")
        $botijaoIndex = Find-HeaderIndex -Headers $headers -Patterns @("*BOTIJ*")
        $receptoraIndex = Find-HeaderIndex -Headers $headers -Patterns @("*RECEPTORA*")
        $obsIndex = Find-HeaderIndex -Headers $headers -Patterns @("*OBS*")
        $serviceIndex = Find-HeaderIndex -Headers $headers -Patterns @("*IATF*", "TE *", "*ACASALAMENTO*")
        $d0Index = Find-HeaderIndex -Headers $headers -Patterns @("D-0*")
        $d8Index = Find-HeaderIndex -Headers $headers -Patterns @("D-8*")
        $rgdIndex = Find-HeaderIndex -Headers $headers -Patterns @("RGD")

        $sheetType = if ($sheet.name -like "*TETF*") { "TETF" } else { "IATF" }
        $roundName = ($sheet.name -replace "º", "a")

        $sheetRecords = @()
        foreach ($row in ($matrix | Select-Object -Skip 1)) {
            if ($row.Count -eq 0 -or $row[0] -notmatch "^0*\d+$") {
                continue
            }

            $dgRaw = if ($dgIndex -ge 0 -and $row.Count -gt $dgIndex) { $row[$dgIndex] } else { "" }
            $serviceRaw = if ($serviceIndex -ge 0 -and $row.Count -gt $serviceIndex) { $row[$serviceIndex] } else { "" }
            $bezerroRaw = if ($bezerroIndex -ge 0 -and $row.Count -gt $bezerroIndex) { $row[$bezerroIndex] } else { "" }
            $d8Raw = if ($d8Index -ge 0 -and $row.Count -gt $d8Index) { $row[$d8Index] } else { "" }
            $dgNormalized = Normalize-Dg -Value $dgRaw
            $dgFechado = -not [string]::IsNullOrWhiteSpace($dgRaw)
            $procedimentoConcluido = $dgFechado
            if ($sheetType -eq "IATF" -and $d8Raw.Trim().ToUpperInvariant() -eq "X") {
                $procedimentoConcluido = $false
            }
            if ($sheetType -eq "TETF" -and $dgNormalized -eq "X") {
                $procedimentoConcluido = $false
            }

            $record = [ordered]@{
                id                = "$($sheet.name)-$($row[0])"
                tecnica           = $sheetType
                rodada            = $roundName
                aba               = $sheet.name
                brinco            = $row[0]
                rgd               = if ($rgdIndex -ge 0 -and $row.Count -gt $rgdIndex) { $row[$rgdIndex] } else { "" }
                d0                = if ($d0Index -ge 0) { Convert-ToDateText -Header $headers[$d0Index] } else { "" }
                d8                = if ($d8Index -ge 0) { Convert-ToDateText -Header $headers[$d8Index] } else { "" }
                statusD8          = $d8Raw
                dataServico       = if ($serviceIndex -ge 0) { Convert-ToDateText -Header $headers[$serviceIndex] } else { "" }
                dataDg            = if ($dgIndex -ge 0) { Convert-ToDateText -Header $headers[$dgIndex] } else { "" }
                servico           = Normalize-Name -Value $serviceRaw
                servicoOriginal   = $serviceRaw
                botijao           = if ($botijaoIndex -ge 0 -and $row.Count -gt $botijaoIndex) { $row[$botijaoIndex] } else { "" }
                dg                = $dgNormalized
                dgOriginal        = $dgRaw
                receptora         = if ($receptoraIndex -ge 0 -and $row.Count -gt $receptoraIndex) { $row[$receptoraIndex] } else { "" }
                observacao        = if ($obsIndex -ge 0 -and $row.Count -gt $obsIndex) { $row[$obsIndex] } else { "" }
                bezerro           = $bezerroRaw
                sexoBezerro       = if ($bezerroRaw -match "M$") { "M" } elseif ($bezerroRaw -match "F$") { "F" } else { "" }
                dgFechado         = $dgFechado
                procedimentoConcluido = $procedimentoConcluido
                prenhez           = $dgNormalized -eq "P+"
                possuiBezerro     = -not [string]::IsNullOrWhiteSpace($bezerroRaw) -and $bezerroRaw -ne "DESCARTE"
            }

            $records += [pscustomobject]$record
            $sheetRecords += [pscustomobject]$record
        }

        $sheetMetas += [pscustomobject][ordered]@{
            aba                    = $sheet.name
            tecnica                = $sheetType
            rodada                 = $roundName
            dimensao               = $worksheet.worksheet.dimension.ref
            linhasPlanilha         = $matrix.Count
            registros              = $sheetRecords.Count
            cabecalho              = $headers
            dataServico            = if ($serviceIndex -ge 0) { Convert-ToDateText -Header $headers[$serviceIndex] } else { "" }
            dataDg                 = if ($dgIndex -ge 0) { Convert-ToDateText -Header $headers[$dgIndex] } else { "" }
            registrosComDgFechado  = @($sheetRecords | Where-Object { $_.dgFechado }).Count
        }
    }

    $closedRecords = @($records | Where-Object { $_.procedimentoConcluido })
    $iatfClosed = @($closedRecords | Where-Object { $_.tecnica -eq "IATF" })
    $tetfClosed = @($closedRecords | Where-Object { $_.tecnica -eq "TETF" })

    function Get-SummaryByTechnique {
        param([object[]]$Items, [string]$Name)

        $positive = @($Items | Where-Object { $_.prenhez }).Count
        [pscustomobject][ordered]@{
            tecnica   = $Name
            total     = $Items.Count
            prenhezes = $positive
            taxa      = Get-Rate -Count $positive -Total $Items.Count
        }
    }

    function Get-GroupedSummary {
        param(
            [object[]]$Items,
            [string]$PropertyName,
            [string]$TotalName,
            [int]$MinTotal = 1
        )

        $Items |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.$PropertyName) } |
            Group-Object -Property $PropertyName |
            ForEach-Object {
                $total = $_.Count
                $positive = @($_.Group | Where-Object { $_.prenhez }).Count
                [pscustomobject][ordered]@{
                    nome      = $_.Name
                    $TotalName = $total
                    prenhezes = $positive
                    taxa      = Get-Rate -Count $positive -Total $total
                }
            } |
            Where-Object { $_.$TotalName -ge $MinTotal } |
            Sort-Object -Property @{ Expression = "taxa"; Descending = $true }, @{ Expression = $TotalName; Descending = $true }, nome
    }

    function Sort-RoundSummaries {
        param([object[]]$Items)

        $Items | Sort-Object -Property @{
            Expression = {
                $match = [regex]::Match($_.nome, "^\d+")
                if ($match.Success) { [int]$match.Value } else { 999 }
            }
        }, @{
            Expression = {
                if ($_.nome -like "*IATF*") { 0 } else { 1 }
            }
        }, nome
    }

    function Get-EffectiveByAnimal {
        param(
            [object[]]$Items,
            [string]$Technique
        )

        $groups = @($Items | Group-Object -Property brinco)
        $animals = @()

        foreach ($group in $groups) {
            $orderedRecords = @($group.Group | Sort-Object -Property rodada)
            $positiveRecords = @($orderedRecords | Where-Object { $_.prenhez })
            $firstPositive = if ($positiveRecords.Count -gt 0) { $positiveRecords[0] } else { $null }
            $rgdValue = @($orderedRecords | Where-Object { $_.rgd } | Select-Object -First 1 -ExpandProperty rgd)

            $animals += [pscustomobject][ordered]@{
                tecnica         = $Technique
                brinco          = $group.Name
                rgd             = if ($rgdValue.Count -gt 0) { $rgdValue[0] } else { "" }
                procedimentos   = $orderedRecords.Count
                prenhezFinal    = $positiveRecords.Count -gt 0
                primeiraPrenhez = if ($null -ne $firstPositive) { $firstPositive.rodada } else { "" }
                ultimoDg        = $orderedRecords[-1].dgOriginal
                ultimoServico   = $orderedRecords[-1].servico
                rodadas         = @($orderedRecords | Select-Object -ExpandProperty rodada)
            }
        }

        $pregnant = @($animals | Where-Object { $_.prenhezFinal }).Count
        $empty = @($animals | Where-Object { -not $_.prenhezFinal }).Count

        return [pscustomobject][ordered]@{
            tecnica             = $Technique
            animaisProtocolados = $animals.Count
            prenhesAoFinal      = $pregnant
            vaziosAoFinal       = $empty
            taxaEfetiva         = Get-Rate -Count $pregnant -Total $animals.Count
            animaisVazios       = @($animals | Where-Object { -not $_.prenhezFinal } | Sort-Object -Property @{ Expression = "procedimentos"; Descending = $true }, brinco)
            distribuicaoPrenhez = @(
                $animals |
                    Where-Object { $_.prenhezFinal } |
                    Group-Object -Property primeiraPrenhez |
                    ForEach-Object {
                        [pscustomobject][ordered]@{
                            rodada  = $_.Name
                            animais = $_.Count
                        }
                    } |
                    Sort-Object -Property rodada
            )
        }
    }

    $byRound = Sort-RoundSummaries -Items (Get-GroupedSummary -Items $closedRecords -PropertyName "rodada" -TotalName "procedimentos")
    $iatfBulls = Get-GroupedSummary -Items $iatfClosed -PropertyName "servico" -TotalName "servicos" -MinTotal 3
    $tetfBulls = Get-GroupedSummary -Items $tetfClosed -PropertyName "servico" -TotalName "transferencias" -MinTotal 2
    $effectiveIatf = Get-EffectiveByAnimal -Items $iatfClosed -Technique "IATF"
    $effectiveTetf = Get-EffectiveByAnimal -Items $tetfClosed -Technique "TETF"
    $effectiveAnimals = @($effectiveIatf.animaisVazios + $effectiveTetf.animaisVazios)

    $repeaters = @(
        @($iatfClosed + $tetfClosed) |
            Group-Object -Property tecnica, brinco |
            ForEach-Object {
                $orderedRecords = @($_.Group | Sort-Object -Property rodada)
                if ($orderedRecords.Count -lt 2) {
                    return
                }

                $positiveRecords = @($orderedRecords | Where-Object { $_.prenhez })
                $firstPositive = if ($positiveRecords.Count -gt 0) { $positiveRecords[0] } else { $null }
                $rgdValue = @($orderedRecords | Where-Object { $_.rgd } | Select-Object -First 1 -ExpandProperty rgd)

                [pscustomobject][ordered]@{
                    tecnica         = $orderedRecords[0].tecnica
                    brinco          = $orderedRecords[0].brinco
                    rgd             = if ($rgdValue.Count -gt 0) { $rgdValue[0] } else { "" }
                    tentativas      = $orderedRecords.Count
                    resultado       = if ($positiveRecords.Count -gt 0) { "Prenhe depois" } else { "Continuou vazio" }
                    primeiraPrenhez = if ($null -ne $firstPositive) { $firstPositive.rodada } else { "" }
                    ultimoDg        = $orderedRecords[-1].dgOriginal
                    rodadas         = @($orderedRecords | Select-Object -ExpandProperty rodada)
                }
            } |
            Sort-Object -Property tecnica, @{ Expression = "tentativas"; Descending = $true }, brinco
    )

    $pregnancyAttemptDistribution = @(
        @($effectiveIatf, $effectiveTetf) |
            ForEach-Object {
                $technique = $_.tecnica
                $_.distribuicaoPrenhez | ForEach-Object {
                    [pscustomobject][ordered]@{
                        tecnica = $technique
                        rodada  = $_.rodada
                        animais = $_.animais
                    }
                }
            }
    )

    $dataAlerts = @(
        [pscustomobject][ordered]@{
            alerta     = "DG pendente ou protocolo descartado"
            quantidade = @($records | Where-Object { -not $_.procedimentoConcluido }).Count
            leitura    = "Registros que nao entraram nos indicadores fechados."
        }
        [pscustomobject][ordered]@{
            alerta     = "Bezerro com DG atual nao prenhe"
            quantidade = @($records | Where-Object { $_.possuiBezerro -and -not $_.prenhez }).Count
            leitura    = "Pode indicar historico de rodada anterior ou preenchimento para revisar."
        }
        [pscustomobject][ordered]@{
            alerta     = "Prenhez sem bezerro informado"
            quantidade = @($records | Where-Object { $_.prenhez -and -not $_.possuiBezerro }).Count
            leitura    = "Acompanhar se o parto ainda nao ocorreu ou se falta atualizar a cria."
        }
    )

    $summary = [pscustomobject][ordered]@{
        generatedAt        = (Get-Date).ToString("s")
        sourceFile         = (Split-Path -Leaf $sourcePath)
        criterio           = "Indicadores calculados apenas com registros numericos e DG preenchido."
        totals             = [ordered]@{
            registrosPlanilha       = $records.Count
            registrosComDgFechado   = $closedRecords.Count
            registrosPendentes      = @($records | Where-Object { -not $_.procedimentoConcluido }).Count
            brincosUnicos           = @($records | Select-Object -ExpandProperty brinco -Unique).Count
        }
        porTecnica         = @(
            Get-SummaryByTechnique -Items $iatfClosed -Name "IATF"
            Get-SummaryByTechnique -Items $tetfClosed -Name "TETF"
        )
        rodadas            = $byRound
        rodadasIatf        = Sort-RoundSummaries -Items (Get-GroupedSummary -Items $iatfClosed -PropertyName "rodada" -TotalName "procedimentos")
        rodadasTetf        = Sort-RoundSummaries -Items (Get-GroupedSummary -Items $tetfClosed -PropertyName "rodada" -TotalName "transferencias")
        tourosIatf         = @($iatfBulls | Select-Object -First 12)
        tourosTetf         = @($tetfBulls | Select-Object -First 12)
        efetividadeAnimal  = @($effectiveIatf, $effectiveTetf)
        repetidoras        = $repeaters
        prenhezPorTentativa = $pregnancyAttemptDistribution
        alertasDados       = $dataAlerts
        abas               = $sheetMetas
        qualidade          = [ordered]@{
            dgPendente                = @($records | Where-Object { -not $_.procedimentoConcluido }).Count
            bezerroComDgNaoPrenhe     = @($records | Where-Object { $_.possuiBezerro -and -not $_.prenhez }).Count
            prenhezSemBezerro         = @($records | Where-Object { $_.prenhez -and -not $_.possuiBezerro }).Count
            grafiasServicoNormalizadas = @("HANCOCK/HANKOK/HANKOCK", "BRUTAO/BRUTÃO ARF", "FENOL/FENOL S", "ARGOS/ARGOS SN")
        }
    }

    $recordsJson = $records | ConvertTo-Json -Depth 8
    $summaryJson = $summary | ConvertTo-Json -Depth 8

    $recordsJson | Set-Content -LiteralPath (Join-Path $outputPath "records.json") -Encoding UTF8
    $summaryJson | Set-Content -LiteralPath (Join-Path $outputPath "summary.json") -Encoding UTF8
    "window.IATF_TETF_SUMMARY = $summaryJson;" | Set-Content -LiteralPath (Join-Path $outputPath "summary.js") -Encoding UTF8
}
finally {
    $zip.Dispose()
}

Write-Host "ETL concluido:"
Write-Host "- data/records.json"
Write-Host "- data/summary.json"
Write-Host "- data/summary.js"
