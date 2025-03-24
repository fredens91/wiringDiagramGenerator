function Extract-ValuesFromExcel {
    # Ottieni il percorso della directory di esecuzione
    $currentPath = (Get-Location).Path
    $xlsxFiles = Get-ChildItem -Path $currentPath -Filter "*.xlsx"

    # Definisci le cartelle di origine e destinazione per i file XML
    $templateXmlFolder = "$currentPath\template XML"
    $assetsFolder = "$currentPath\assets"

    # Creazione dell'oggetto Excel
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false

    # Rimuove tutti i file già presenti nella cartella "assets" prima di iniziare
    if (Test-Path $assetsFolder) {
        Get-ChildItem -Path $assetsFolder -File | Remove-Item -Force
    } else {
        # Crea la cartella "assets" se non esiste
        New-Item -Path $assetsFolder -ItemType Directory
    }

    foreach ($file in $xlsxFiles) {
        try {
            # Apre il file Excel
            $workbook = $excelApp.Workbooks.Open($file.FullName)

            # Ottieni il primo foglio di lavoro (se ce n'è più di uno, seleziona quello che ti serve)
            $worksheet = $workbook.Sheets.Item(1)

            # Itera attraverso le righe a partire dalla riga 6
            $row = 6
            while ($worksheet.Cells.Item($row, 4).Value2 -ne $null -and $worksheet.Cells.Item($row, 6).Value2 -ne $null) {
                $valueD = $worksheet.Cells.Item($row, 4).Value2
                $valueF = $worksheet.Cells.Item($row, 6).Value2

                if ($valueD -ne $null -and $valueD -ne '' -and $valueF -ne $null -and $valueF -ne '') {
                    # Stampa la coppia di valori "valueD: valueF" su una nuova riga
                    Write-Host "$($valueD): $($valueF)"

                    # Copia i file XML dalla cartella "template XML" nella cartella "assets"
                    $xmlFileName = "$templateXmlFolder\$valueD.xml"

                    if (Test-Path $xmlFileName) {
                        # Copia il file XML n volte nella cartella "assets"
                        for ($i = 1; $i -le $valueF; $i++) {
                            $destinationPath = "$assetsFolder\$valueD-$i.xml"
                            Copy-Item -Path $xmlFileName -Destination $destinationPath
                            Write-Host "File copiato a: $destinationPath"
                        }
                    } else {
                        Write-Host "File XML per $valueD non trovato nella cartella 'template XML'"
                    }
                }
                $row++
            }

            # Chiude il workbook senza salvare
            $workbook.Close($false)
        } catch {
            Write-Host "Errore durante l'elaborazione del file $($file.Name): $_"
        }
    }

    # Esci dall'applicazione Excel
    $excelApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp)
}

# Chiamata alla funzione
Extract-ValuesFromExcel
