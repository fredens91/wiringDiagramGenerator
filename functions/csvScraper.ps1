# Path della cartella che contiene i file .csv
$folderPath = "C:\Users\orlandi_f\Downloads\"

# Ottieni tutti i file CSV nella cartella
$csvFiles = Get-ChildItem -Path $folderPath -Filter "*.csv"

# Per ogni file CSV trovato, esegui il parsing e la creazione dell'XML
ForEach ($csvFile in $csvFiles) {
    # Path del file CSV corrente
    $csvPath = $csvFile.FullName

    # Path del file XML di output (utilizza lo stesso nome del file CSV ma con estensione .xml)
    $xmlPath = [System.IO.Path]::ChangeExtension($csvPath, ".xml")

    # Carica il CSV e filtra le righe vuote
    $csvData = Import-Csv -Path $csvPath -Delimiter ";" | Where-Object { $_ -ne $null -and $_.'Nome' -ne "" }

    # Debug: stampa i dati del CSV per controllare
    Write-Host "Contenuto del CSV:"
    $csvData | Format-Table -AutoSize

    # Verifica i nomi delle colonne
    Write-Host "Intestazioni delle colonne nel CSV:"
    $csvData[0].PSObject.Properties.Name

    # Assumiamo che la colonna "Nome" si chiami esattamente "Nome", altrimenti aggiorna qui
    # Estrai i dati dalla colonna "Nome" partendo dalla riga 6
    $colonnaNome = $csvData[5..($csvData.Length - 1)] | Where-Object { $_.'Nome' -ne "" } | ForEach-Object { $_.'Nome' }

    # Stampa i valori trovati per la colonna "Nome" per debug
    Write-Host "Valori trovati nella colonna 'Nome':"
    $colonnaNome

    # Crea l'oggetto XML
    [xml]$xml = New-Object System.Xml.XmlDocument

    # Crea l'elemento radice
    $root = $xml.CreateElement("Items")

    # Aggiungi gli elementi per ogni valore trovato nella colonna "Nome"
    foreach ($value in $colonnaNome) {
        $item = $xml.CreateElement("Item")
        $item.InnerText = $value
        $root.AppendChild($item)
    }

    # Aggiungi l'elemento radice al documento XML
    $xml.AppendChild($root)

    # Salva l'XML su disco
    $xml.Save($xmlPath)

    Write-Host "File XML creato con successo per $csvPath! Il file si trova in: $xmlPath"
}
