# Funzione principale che unisce entrambe le operazioni
function Main {
    param (
        [string]$folderPath
    )

    # Ottieni il percorso assoluto della cartella dove si trova il file script (main.ps1)
    $scriptDir = [System.IO.Path]::GetDirectoryName((Get-Command $MyInvocation.MyCommand.Name).Source)

    # Crea i percorsi relativi alla root dello script
    $xlsxFolderPath = $scriptDir # La cartella dove si trova lo script contiene i file .xlsx
    $templateFolder = Join-Path -Path $scriptDir -ChildPath "template XML"  # Cartella template XML
    $assetsFolder = Join-Path -Path $scriptDir -ChildPath "assets"  # Cartella assets

    # Verifica se la cartella template XML esiste
    if (-Not (Test-Path -Path $templateFolder)) {
        Write-Host "Il percorso template XML non esiste."
        return
    }

    # Ottieni tutti i file .xlsx nella cartella
    $xlsxFiles = Get-ChildItem -Path $xlsxFolderPath -Filter "*.xlsx"

    # Per ogni file trovato, esegui la conversione da XLSX a CSV
    foreach ($file in $xlsxFiles) {
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

        Write-Host "Conversione da XLSX a CSV completata! Il file CSV si trova in: $csvPath"
    }

    # Funzione per convertire CSV in XML
    function Convert-CsvToXml {
        param (
            [string]$path
        )

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

            Write-Host "File CSV convertito in XML: $xmlFilePath"
        }
    }

    # Esegui la conversione da CSV a XML
    Convert-CsvToXml -path $folderPath

    # Funzione per cercare e copiare i file XML in base a stringhe comuni
    function Search-AndCopyFiles {
        # Ottieni tutti i file .xml nella cartella "template XML"
        $templateFiles = Get-ChildItem -Path $templateFolder -Filter "*.xml"

        # Ottieni tutti i file .xml nella cartella dello script
        $downloadsFiles = Get-ChildItem -Path $folderPath -Filter "*.xml"

        # Verifica se la cartella "assets" esiste, altrimenti creala
        if (-not (Test-Path -Path $assetsFolder)) {
            New-Item -Path $assetsFolder -ItemType Directory
            Write-Host "La cartella 'assets' è stata creata in: $assetsFolder"
        }

        # Per ogni file nella cartella "template XML"
        foreach ($templateFile in $templateFiles) {
            # Estrai il nome del file senza estensione
            $templateFileName = [System.IO.Path]::GetFileNameWithoutExtension($templateFile.Name)

            # Per ogni file .xml nella cartella dello script
            foreach ($downloadsFile in $downloadsFiles) {
                # Leggi il contenuto del file XML
                $xmlContent = Get-Content -Path $downloadsFile.FullName -Raw

                # Cerca le stringhe comuni di almeno 4 caratteri tra il nome del file e il contenuto del file XML
                $matches = [System.Text.RegularExpressions.Regex]::Matches($xmlContent, "\w{4,}")

                # Verifica se una delle stringhe trovate è uguale al nome del file
                foreach ($match in $matches) {
                    if ($match.Value -eq $templateFileName) {
                        # Se c'è una corrispondenza, copia il file .xml nella cartella "assets"
                        $destinationPath = [System.IO.Path]::Combine($assetsFolder, $templateFile.Name)
                        Copy-Item -Path $templateFile.FullName -Destination $destinationPath
                        Write-Host "File copiato: $destinationPath"
                        break
                    }
                }
            }
        }
    }

    # Esegui la funzione di ricerca e copia
    Search-AndCopyFiles
}

# Esegui la funzione principale con il percorso specificato
Main -folderPath $scriptDir
