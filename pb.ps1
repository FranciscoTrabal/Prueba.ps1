$excelProcesses = Get-Process -Name "excel" -ErrorAction SilentlyContinue

# Detener todos los procesos de Excel en ejecución
if ($excelProcesses) {
    foreach ($process in $excelProcesses) {
        Stop-Process -Id $process.Id -Force
    }
}

Copy-Item "C:\Program Files\Software\_s2f02t3\T2ra0b2al3.mdb" -Destination "C:\DatosSAF\SAF_CONSULTAS\TRABAL_2022_B.mdb" -Force
Copy-Item "C:\Program Files\Software\_s2f02t3\T2ra0b2al3.mdb" -Destination "C:\DatosSAF\SAF_CONSULTAS\TRABAL_2022.mdb" -Force
Copy-Item "C:\Program Files\Software\_s2f02t3\T2ra0b2al3.mdb" -Destination "\\10.0.0.61\saf_data_reporting\TRABAL_2022.mdb" -Force
###############################################################################################################################

$ORIGEN ="C:\DatosSAF\SAF_CONSULTAS\TRABAL_2022_B.mdb"
$destino4 = "C:\DatosSAF\DB\BASEdatos\TRABAL_2022.mdb"

#Copy-Item $ORIGEN -Destination $DESTINO3 -Force
Copy-Item $ORIGEN -Destination $destino4 -Force


$Fecha   = get-date -uformat %d%m%Y%H%M%S
$Nombre  ="TRABAL_2022_"
$Tipo    =".zip"
$Ruta    ="C:\Users\Administrator\Documents\Back\OneDrive\Back\"
$Archivo = $Nombre+$Fecha+$Tipo
$Destino =$Ruta+$Archivo

Compress-Archive -LiteralPath $destino4 -CompressionLevel Optimal -DestinationPath $Destino 
#Compress-Archive -LiteralPath 'C:\Users\Administrator\Documents\saf\TRABAL_2022.mdb' -CompressionLevel Optimal -DestinationPath $Destino

if($numArchivos == 6){
    
    Install-Module -Name Microsoft.Graph.OneDriveFiles -Scope CurrentUser -Force -AllowClobber
    Import-Module Microsoft.Graph.OneDriveFiles

    # Inicia sesión en tu cuenta de OneDrive
    Connect-MgGraph -Scopes "Files.ReadWrite.All"

    # Especifica la ruta del archivo que deseas eliminar
    $filePath = "/ruta/del/archivo/archivo.txt"

    # Obtén la instancia del archivo que deseas eliminar
    $file = Get-MgUserOneDriveFile -Path $filePath

    # Elimina el archivo
    Remove-MgUserOneDriveFile -Item $file

}