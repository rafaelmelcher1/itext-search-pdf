Add-Type -Path "C:\itextsharp\lib\itextsharp.dll"

# Diretorio onde os arquivos PDF estao armazenados
$dir = "F:\"

# Palavra a ser pesquisada
$palavra = "name"

# Array para armazenar os resultados
$results = @()

# Loop pelos arquivos PDF no diretorio e subpastas
Get-ChildItem -Path $dir -Filter *.pdf -Recurse | ForEach-Object {
    $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $_.FullName
    for ($page = 1; $page -le $reader.NumberOfPages; $page++) {
        $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page)
        if ($text.Contains($palavra)) {
            $result = New-Object PSObject -Property @{
                FullName = $_.FullName
                Directory = $_.Directory.Name
                Name = $_.Name
                Page = $page
            }
            $results += $result
        }
    }
    $reader.Close()
}

# Exibir os resultados em uma tabela
$results | Select-Object FullName, Directory, Name, Page | Format-Table -AutoSize