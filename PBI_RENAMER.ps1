param(
    [string] $sourceFileNm, 
    [string] $newConnection
)

$file = "Mod_"+$sourceFileNm
copy-Item $sourceFileNm".pbix" -Destination $file".pbix"

# Change  input.pbix file to zip
Rename-Item -Path $file".pbix" -newname $file".zip"

# Unzip
$zip = $file+'.zip'
# echo $file
# echo $zip
Expand-Archive -Path $zip

# Delete old zip 
remove-Item $zip

# Replace  Connections string
$regexRule = '\Data Source=(.*?)\;'
$inputFile = $file+'\Connections'
(Get-Content $inputFile)| Foreach-Object {$_ -replace $regexRule,-join('Data Source=',$newConnection,';')} | Out-File $inputFile -NoNewline -Encoding ASCII

# Zip changed files
Compress-7Zip -Path $file -ArchiveFileName $file".zip" -Format Zip -CompressionLevel Normal

# Rename new modifed file
Rename-Item -Path $file".zip" -newname $file".pbix"
