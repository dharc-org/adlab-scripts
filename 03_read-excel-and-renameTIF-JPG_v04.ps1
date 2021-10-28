# 03_read-excel-and-renameTIF-JPG
# v04

#User input can be read like this:
#$num = Read-Host "Store number"

# prendo il nome del File Excel
# lo utilizzo per rinominare la prima parte del file immagine
# prendo i dati presenti all'interno per aggiungere in fondo il label sia dei Tif che dei JPG
# rinomino anche la cartella

###### DA PERSONALIZZARE ######
$ImagePrefixExt = "nome-immagine-esistente"
$ImagePrefixNew = "BO0451_CAM9487"

$FolderPathTif = "E:\SCANSIONI\folder-name"
$FolderPathJpg = "E:\SCANSIONI\folder-name\folder-name_jpg"
##############################


$FolderRoot = (Split-Path -Path $FolderPathTif) + "\"
$FolderPathTif = $FolderPathTif.TrimEnd('\')+"\"
$FolderPathJpg = $FolderPathJpg.TrimEnd('\')+"\"
$ExcelPath = $FolderPathTif+$ImagePrefixNew+".xlsx"

if (-not(Test-Path -Path $ExcelPath -PathType leaf)) {
	Write-Host "Non ho trovato il file Excel:"$ImagePrefixNew".xlsx"
	exit
}


Write-Host "1) Read content of Excel file: " + $ExcelPath
#Open Excel
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
$Workbook = $Excel.Workbooks.Open($ExcelPath)
$Sheet = $Workbook.ActiveSheet
$UsedRange = $Sheet.UsedRange
$RowMax = ($Sheet.UsedRange.Rows).count
$ColMax = ($Sheet.UsedRange.Columns).count

#Loop inside the Sheet
$ArrayLabel = @("array")
for ($i = 1; $i -le $RowMax; $i++) {
	Write-Host $Sheet.Cells.Item($i, 1).Text
	$ArrayLabel += $Sheet.Cells.Item($i, 1).Text
}


Write-Host "2) Rename TIF files from: " + $FolderPathTif
$i = 1
Get-ChildItem -Path $FolderPathTif -Filter *.tif |
Sort-Object | ForEach-Object {
   $extension = $_.Extension
   $newName = ($_.BaseName -replace $ImagePrefixExt,$ImagePrefixNew) + "_" + $ArrayLabel[$i] + $_.Extension
   Rename-Item -Path $_.FullName -NewName $newName 
   $i++
}

Write-Host "3) Rename JPG files from: " + $FolderPathJpg
$i = 1
Get-ChildItem -Path $FolderPathJpg -Filter *.jpg |
Sort-Object | ForEach-Object {
   $extension = $_.Extension
   $newName = ($_.BaseName -replace $ImagePrefixExt,$ImagePrefixNew) + "_" + $ArrayLabel[$i] + $_.Extension
   Rename-Item -Path $_.FullName -NewName $newName 
   $i++
}


#Clean up after you're done:
$Workbook.Close()
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)


#Rename Folders
$FolderPathJpgNew = $ImagePrefixNew+"_jpg"
$FolderPathTifNew = $ImagePrefixNew

Write-Host "4) Rename Folders " 

Write-Host "da "$FolderPathJpg" a "$FolderPathJpgNew
Rename-Item -Path $FolderPathJpg -NewName $FolderPathJpgNew

Write-Host "da "$FolderPathTif" a "$FolderPathTifNew
Rename-Item -Path $FolderPathTif -NewName $FolderPathTifNew
