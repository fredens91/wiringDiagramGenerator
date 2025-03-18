function Search-AndCopyFiles {
    # Cartelle di origine e destinazione
    $templateFolder = "C:\Users\orlandi_f\Downloads\template XML"
    $downloadsFolder = "C:\Users\orlandi_f\Downloads"
    $assetsFolder = "C:\Users\orlandi_f\Downloads\assets"

    # Ottieni tutti i file .xml nella cartella "template XML"
    $templateFiles = Get-ChildItem -Path $templateFolder -Filter "*.xml"

    # Ottieni tutti i file .xml nella cartella "Downloads"
    $downloadsFiles = Get-ChildItem -Path $downloadsFolder -Filter "*.xml"

    # Crea la cartella "assets" se non esiste
    if (-not (Test-Path -Path $assetsFolder)) {
        New-Item -Path $assetsFolder -ItemType Directory
    }

    # Per ogni file nella cartella "template XML"
    foreach ($templateFile in $templateFiles) {
        # Estrai il nome del file senza estensione
        $templateFileName = [System.IO.Path]::GetFileNameWithoutExtension($templateFile.Name)

        # Per ogni file .xml nella cartella "Downloads"
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

# Esegui la funzione
Search-AndCopyFiles
