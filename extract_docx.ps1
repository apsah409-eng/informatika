param($filename)
Copy-Item $filename -Destination 'temp.zip'
Expand-Archive -Path 'temp.zip' -DestinationPath 'temp_docx' -Force
[xml]$doc = Get-Content 'temp_docx\word\document.xml'
$text = $doc.document.body.p | ForEach-Object { $_.innerText } | Where-Object { $_ -ne $null -and $_ -match '\S' }
$text | Out-File "extracted.txt" -Encoding utf8
Remove-Item -Recurse -Force 'temp_docx'
Remove-Item 'temp.zip'
