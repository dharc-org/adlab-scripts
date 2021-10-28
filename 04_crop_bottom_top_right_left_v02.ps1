# 04_crop_bottom_top_right_left 
# v0.2

###### DA PERSONALIZZARE ######
$FolderJpg  = "E:\SCANSIONI\folder-name\folder-name_jpg"

$CropTop = 365
$CropBottom = 380
$CropRight = 345
$CropLeft = 345
####################################

$FolderJpg = $FolderJpg.TrimEnd('\')+"\"
Write-Host 'FolderJpg: '$FolderJpg

write-host(" ")
write-host("Script CROP from Top (North) pixel "+$CropTop)
write-host("Script CROP from Bottom (South) pixel "+$CropBottom)
write-host("Crop ODD Left (East) pixel "+$CropLeft)
write-host("Crop EVEN Right (West) pixel "+$CropRight)
$x = 1;
Get-ChildItem $FolderJpg -Filter *.jpg | 
Sort-Object | Foreach-Object {
	if($x % 2 -eq 0 ){
     	Write-Host(". EVEN "+$x+" crop Right: "+$_.FullName)
		#magick.exe $_.FullName -gravity East -chop '$CropRight'x0 ($_.FullName)
		magick.exe $_.FullName -crop +0+$CropTop -crop -$CropRight-$CropBottom ($_.FullName)
	} else {
		Write-Host(". ODD  "+$x+" crop Left: "+$_.FullName)
		#magick.exe $_.FullName -gravity West -chop 0x$CropLeft ($_.FullName)
		magick.exe $_.FullName -crop +$CropLeft+$CropTop -crop -0-$CropBottom ($_.FullName)
	}
	$x++;
}
