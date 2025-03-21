function Extract-ValuesFromExcel {
    # Ottieni il percorso della directory di esecuzione
    $currentPath = (Get-Location).Path
    $xlsxFiles = Get-ChildItem -Path $currentPath -Filter "*.xlsx"

    # Creazione dell'oggetto Excel
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false

    foreach ($file in $xlsxFiles) {
        try {
            # Apre il file Excel
            $workbook = $excelApp.Workbooks.Open($file.FullName)

            # Ottieni il primo foglio di lavoro (se ce n'è più di uno, seleziona quello che ti serve)
            $worksheet = $workbook.Sheets.Item(1)

            # Crea una lista per memorizzare le coppie di valori D e F
            $pairs = @()

            # Itera attraverso le righe a partire dalla riga 6
            $row = 6
            while ($worksheet.Cells.Item($row, 4).Value2 -ne $null -and $worksheet.Cells.Item($row, 6).Value2 -ne $null) {
                $valueD = $worksheet.Cells.Item($row, 4).Value2
                $valueF = $worksheet.Cells.Item($row, 6).Value2

                if ($valueD -ne $null -and $valueD -ne '' -and $valueF -ne $null -and $valueF -ne '') {
                    # Aggiungi la coppia di valori a $pairs
                    $pairs += "$valueD;$valueF"
                }
                $row++
            }

            # Stampa i valori separati da ";"
            $pairs -join ';' | Write-Host

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
