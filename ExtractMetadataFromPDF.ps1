# Define the input and output file paths
$pdfFilePath = ".\Repertorium_burst_def_-_ADRESSEN_WZC-_PROVINCIE_Vlaams-Brabant.pdf"

$textFilePath = ".\pdftotextfile.txt"
$outputFilePath = ".\output.csv"

# Extract text from PDF using pdftotext: https://www.xpdfreader.com/download.html -> Download the Xpdf command line tools
Start-Process -FilePath "<path-to>\xpdf-tools-win-4.05\xpdf-tools-win-4.05\bin64\pdftotext.exe" -ArgumentList "$pdfFilePath $textFilePath" -Wait

# Read the content of the extracted text file
$fileContent = Get-Content -Path $textFilePath

# Initialize variables to hold the extracted data
$dossiernr = ""
$erkenningsnr = ""
$name = ""
$tel = ""
$email = ""
$url = ""
$data = @()

# Loop through each line in the file content
# Validate regex with https://regex101.com/
foreach ($line in $fileContent) {
    if ($line -match "Dossiernr") {
        if ($line -match "Dossiernr\.: ([^\s]+)") {
            $dossiernr = $matches[1].Trim()
        }
        if ($line -match "Erkenningsnr\.: ([^\s]+) ([^\d]+)") {
            $erkenningsnr = $matches[1].Trim()
            $name = $matches[2].Trim()
        }
        if ($line -match "tel\. : (.+?) e-mail:") {
            $tel = $matches[1].Trim()
        }
        if ($line -match "e-mail: ([^\s]+) ([^\d]+)") {
            $email = $matches[1].Trim()
        }
        if ($line -match "url: ([^\s]+) ([^\d]+)") {
            $url = $matches[1].Trim()
        }

        $data += [PSCustomObject]@{
            Dossiernr = $dossiernr
            Erkenningsnr = $erkenningsnr
            Name = $name
            Tel = $tel
            Email = $email
            Url = $url
        }
        # Reset variables for the next company
        $dossiernr = ""
        $erkenningsnr = ""
        $name = ""
        $tel = ""
        $email = ""
        $url = ""
        
    }
}

# Write the data to a CSV file
$data | Export-Csv -Path $outputFilePath -NoTypeInformation

Write-Output "Data extraction complete. Output saved to $outputFilePath"
