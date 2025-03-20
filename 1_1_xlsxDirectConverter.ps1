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

            # Crea una lista per memorizzare i valori non nulli della colonna D
            $values = @()

            # Itera attraverso le righe della colonna D a partire dalla riga 6
            $row = 6
            while ($worksheet.Cells.Item($row, 4).Value2 -ne $null) {
                $value = $worksheet.Cells.Item($row, 4).Value2
                if ($value -ne $null -and $value -ne '') {
                    $values += $value
                }
                $row++
            }

            # Crea il percorso del file di output .txt
            $txtFilePath = [System.IO.Path]::Combine($currentPath, "Valori_Estratti.txt")

            # Scrivi i valori in un file .txt separati da ";"
            $values -join ';' | Out-File -FilePath $txtFilePath -Encoding UTF8

            Write-Host "I valori sono stati estratti in: $txtFilePath"

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
