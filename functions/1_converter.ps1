function Convert-XlsxToCsv {
    # Ottieni il percorso della directory dello script
    $currentPath = (Get-Location).Path
    $xlsxFiles = Get-ChildItem -Path $currentPath -Filter "*.xlsx"

    # Creazione dell'oggetto Excel
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false

    foreach ($file in $xlsxFiles) {
        try {
            # Apre il file Excel
            $workbook = $excelApp.Workbooks.Open($file.FullName)

            # Percorso del file CSV da generare
            $csvPath = [System.IO.Path]::Combine($currentPath, "$($file.BaseName).csv")
            
            # Salva il contenuto del foglio come CSV
            $workbook.SaveAs($csvPath, 6)  # 6 è il formato per CSV

            # Chiude il workbook
            $workbook.Close($false)
        } catch {
            Write-Host "Errore durante la conversione del file $($file.Name): $_"
        }
    }

    # Esci dall'applicazione Excel
    $excelApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp)
}

# Chiamata alla funzione
Convert-XlsxToCsv