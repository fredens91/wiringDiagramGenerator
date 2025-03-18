# Path della cartella che contiene i file .xlsx
$folderPath = "C:\Users\orlandi_f\Downloads\"

# Ottieni tutti i file .xlsx nella cartella
$files = Get-ChildItem -Path $folderPath -Filter "*.xlsx"

# Per ogni file trovato, esegui la conversione
ForEach ($file in $files) {
    # Imposta il percorso del file di input e output
    $xlsxPath = $file.FullName
    $csvPath = [System.IO.Path]::ChangeExtension($xlsxPath, ".csv")

    # Avvia una nuova istanza di Excel (in background)
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false  # Imposta a $true se vuoi vedere Excel aprirsi

    # Apre il file Excel
    $Workbook = $Excel.Workbooks.Open($xlsxPath)

    # Salva il workbook come CSV
    $Workbook.SaveAs($csvPath, 6)  # 6 è il codice per CSV

    # Chiudi il file senza salvare
    $Workbook.Close($false)

    # Rilascia l'oggetto Excel
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
    $Excel = $null

    Write-Host "Conversione completata! Il file CSV si trova in: $csvPath"
}
