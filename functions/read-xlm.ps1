# Funzione per convertire CSV in XML
function Convert-CsvToXml {
    param (
        [string]$path
    )

    # Verifica se il percorso esiste
    if (Test-Path -Path $path) {
        # Ottieni tutti i file CSV nel percorso
        $csvFiles = Get-ChildItem -Path $path -Filter *.csv
        
        # Itera su ciascun file CSV
        foreach ($csvFile in $csvFiles) {
            # Leggi il contenuto del file CSV
            $csvContent = Import-Csv -Path $csvFile.FullName

            # Converti il contenuto in XML
            $xmlContent = $csvContent | ConvertTo-Xml -As String -Depth 3

            # Crea il percorso per il file XML di output
            $xmlFilePath = [System.IO.Path]::Combine($path, [System.IO.Path]::GetFileNameWithoutExtension($csvFile.Name) + ".xml")

            # Salva il contenuto XML in un file
            $xmlContent | Out-File -FilePath $xmlFilePath

            Write-Host "File convertito: $xmlFilePath"
        }
    } else {
        Write-Host "Il percorso specificato non esiste."
    }
}

# Esegui la funzione con il percorso specificato
Convert-CsvToXml -path "C:\Users\orlandi_f\Downloads\"
