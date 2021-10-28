# 02_convert-to-jpg 
# v0.7

###### DA PERSONALIZZARE ######
$folderTif = "E:\SCANSIONI\folder-name"
###############################

$folderTif = $folderTif.TrimEnd('\')+"\"

Write-Host '$FolderTif: '$folderTif
$dirName = Split-Path $folderTif -Leaf
Write-Host '$dirName: '$dirName
$folderJpg = $folderTif+$dirName+'_jpg'
Write-Host '$FolderJpg: '$folderJpg

mkdir $folderJpg

Write-Host ' '

Get-ChildItem $folderTif -Filter *.TIF | 
Sort-Object | Foreach-Object {
    Write-Host '.convert file:' + $_.BaseName + ".tif";
    magick.exe -units PixelsPerInch $_.FullName -density 150 -resize 2048x -interlace JPEG -colorspace RGB -sampling-factor 4:2:0 -strip -quality 70% -set filename: "%t" $folderJpg/%[filename:].jpg
 }
 
