# Assicurati di avere il modulo ImportExcel installato:
# Install-Module -Name ImportExcel -Force

$path = "C:\Users\orlandi_f\Documents\GitHub\wiringDiagramGenerator\template XML"
$outputXlsx = "C:\Users\orlandi_f\Documents\GitHub\wiringDiagramGenerator\folders_list.xlsx"

# Ottieni l'elenco delle sottocartelle
$folders = Get-ChildItem -Path $path -Directory

# Creare un array di oggetti per inserire in Excel
$excelData = @()

# Cicla attraverso ogni sottocartella
foreach ($folder in $folders) {
    $bomPath = Join-Path $folder.FullName "BOM"
    
    # Inizializza una mappa per i dati della cartella corrente
    $row = @{"NomeCartella" = $folder.Name}

    if (Test-Path $bomPath -PathType Container) {
        # Se la cartella BOM esiste, cerca il file XLSX all'interno
        $bomFile = Get-ChildItem -Path $bomPath -Filter "*.xlsx" | Select-Object -First 1

        if ($bomFile) {
            # Carica il file Excel
            $bomData = Import-Excel -Path $bomFile.FullName

            # Verifica se c'è una colonna D e se contiene dati
            if ($bomData[0].PSObject.Properties["D"]) {
                # Estrai i valori non nulli dalla colonna D a partire dalla riga 5
                $values = $bomData | Where-Object { $_.D -ne $null } | Select-Object -ExpandProperty D

                # Aggiungi i valori estratti alle colonne successive
                $columnIndex = 1  # Comincia dalla colonna B (indice 1)

                foreach ($value in $values) {
                    $row["Colonna$columnIndex"] = $value
                    $columnIndex++
                }
            } else {
                Write-Host "La colonna D non esiste nel file $($bomFile.Name) della cartella $($folder.Name)."
            }
        } else {
            Write-Host "Nessun file XLSX trovato nella cartella BOM di $($folder.Name)."
        }
    }

    # Aggiungi la riga (con o senza dati BOM) all'array
    $excelData += New-Object PSObject -Property $row
}

# Scrivi i dati in un file XLSX
$excelData | Export-Excel -Path $outputXlsx -WorksheetName "Sottocartelle" -AutoSize -TableName "FoldersList"

Write-Host "Elenco sottocartelle con valori estratti salvato in $outputXlsx"
