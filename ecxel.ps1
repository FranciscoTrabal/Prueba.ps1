$excelProcesses = Get-Process -Name "excel" -ErrorAction SilentlyContinue

# Detener todos los procesos de Excel en ejecución
if ($excelProcesses) {
    foreach ($process in $excelProcesses) {
        Stop-Process -Id $process.Id -Force
    }
}
#Copy-Item "C:\Program Files\Software\_s2f02t3\T2ra0b2al3.mdb" -Destination "C:\DatosSAF\SAF_CONSULTAS\TRABAL_2022_B.mdb" -Force
#
Copy-Item "C:\Users\DESARROLLO\Desktop\PB\pb.txt" -Destination "C:\Users\DESARROLLO\Desktop\pb.txt" -Force

#$onedriveall = "C:\Users\Administrator\Documents\Back\OneDrive\Back\*"
#$onedrive = "C:\Users\Administrator\Documents\Back\OneDrive\Back\"
#
$onedrive = "C:\Users\DESARROLLO\Desktop\pone\"
$onedriveall = "C:\Users\DESARROLLO\Desktop\pone\*"
$original = "C:\Users\DESARROLLO\Desktop\pz"
$ORIGEN ="C:\Users\DESARROLLO\Desktop\PB\pb.txt"
$destino4 = "C:\Users\DESARROLLO\Desktop\pb.txt"

#Copy-Item $ORIGEN -Destination $DESTINO3 -Force
Copy-Item $ORIGEN -Destination $destino4 -Force


$Fecha   = get-date -uformat %d%m%Y%H%M%S
$Nombre  ="Texo_"
$Tipo    =".zip"
$Ruta    ="C:\Users\DESARROLLO\Desktop\pz\" #Debo cambiar esta ruta por las onedrive que al parecer es "C:\Users\Administrator\Documents\Back\OneDrive\Back\"
$Archivo = $Nombre+$Fecha+$Tipo

$Destino =$Ruta+$Archivo

Compress-Archive -LiteralPath $destino4 -CompressionLevel Optimal -DestinationPath $Destino 

$numArchivos =  (Get-ChildItem $original -Recurse | Measure-Object).Count #Debo cambiar esta ruta por las onedrive que al parecer es "C:\Users\Administrator\Documents\Back\OneDrive\Back\"


if($numArchivos -eq 6){
    $original = $original+"\*"
    Remove-Item $onedriveall #Aqui debemos remover los zips y subir las copias  #Debo cambiar esta ruta por las onedrive que al parecer es "C:\Users\Administrator\Documents\Back\OneDrive\Back\"
    Copy-Item $original -Destination $onedrive -Force
    Remove-Item $original
}

