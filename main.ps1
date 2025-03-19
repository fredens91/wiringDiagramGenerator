function Convert-ExcelToXml {
    param (
        [string]$XmlFileName = "output.xml"
    )

    # Ottenere la directory dello script
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    
    # Trovare il primo file .xlsx nella root
    $ExcelFile = Get-ChildItem -Path $scriptRoot -Filter "*.xlsx" | Select-Object -First 1

    if (-not $ExcelFile) {
        Write-Error "❌ Nessun file Excel trovato nella directory: $scriptRoot"
        return
    }

    $ExcelFilePath = Join-Path -Path $scriptRoot -ChildPath $ExcelFile.Name
    $XmlFilePath = Join-Path -Path $scriptRoot -ChildPath $XmlFileName

    # Importa i dati Excel
    $excelData = Import-Excel -Path $ExcelFilePath -Worksheet 1

    if (-not $excelData) {
        Write-Error "❌ Nessun dato trovato nel file Excel."
        return
    }

    # Estrarre i dati dalle colonne D e F (Codice e Quantità)
    $codici = @()
    $quantita = @()

    for ($i = 5; $true; $i++) {
        $codice = $excelData | Select-Object -ExpandProperty "D$i" -ErrorAction SilentlyContinue
        $qta = $excelData | Select-Object -ExpandProperty "F$i" -ErrorAction SilentlyContinue
        
        if (-not $codice -and -not $qta) { break } # Interrompe quando entrambe sono vuote

        if ($codice -and $qta) {
            $codici += $codice
            $quantita += $qta
        }
    }

    if ($codici.Count -eq 0) {
        Write-Error "❌ Nessun dato valido trovato nelle colonne Codice e Quantità."
        return
    }

    # Creazione del documento XML
    $xmlDoc = New-Object System.Xml.XmlDocument
    $root = $xmlDoc.CreateElement("Dati")
    $xmlDoc.AppendChild($root)

    for ($i = 0; $i -lt $codici.Count; $i++) {
        $item = $xmlDoc.CreateElement("Elemento")

        $codiceNode = $xmlDoc.CreateElement("Codice")
        $codiceNode.InnerText = $codici[$i]
        $item.AppendChild($codiceNode)

        $quantitaNode = $xmlDoc.CreateElement("Quantità")
        $quantitaNode.InnerText = $quantita[$i]
        $item.AppendChild($quantitaNode)

        $root.AppendChild($item)
    }

    # Salvataggio XML
    $xmlDoc.Save($XmlFilePath)
    Write-Output "✅ File XML creato: $XmlFilePath"
}
