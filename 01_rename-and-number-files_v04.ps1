# 01_rename-and-number-files
# v04

# input argument: Extension
# i.e. .\rename-and-number-files_v04.psi jpg

###### DA PERSONALIZZARE ######
$folder = "E:\SCANSIONI\folder-name"
$name = "Nome-Opera"
###############################


if ($args[0] -eq $null) {
	Write-Host "Lo script ha bisogno di un argomento, ad esempio: `n .\rename-and-number-files_v04.ps1 jpg"
	exit
}

$extension = "."+$args[0]
$folder = $Folder.TrimEnd('\')+"\"

$findExt = "*"+$extension
$files = Get-ChildItem $folder -Filter $findExt | Sort-Object
$x = 1
foreach ($file in $files) 
{
	$newName=$folder+$name+"_"+$x.ToString("0000")+$extension
	$filePath=$folder+$file
	Rename-Item $filePath $newName
	$x++
	#Write-Host $Folder" file: "$file" - newName: "$newName " -> " $filePath
}

$x--

Write-Host "Rinominato" $x "immagini" $extension "dalla cartella" $folder