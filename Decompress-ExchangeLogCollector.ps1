[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$Directory = [Microsoft.VisualBasic.Interaction]::InputBox('Informe o diretorio onde o dados serao descompactados. O diretório não pode ser uma subpasta. Siga o Exemplo:', 'Decompress Path', "D:\ExchangeLogCollector")

$result = [System.Windows.Forms.MessageBox]::Show('O Script ExchangeLogCollector foi executado a partir de um Servidor Exchange?' , "Tipo de Coleta" , 4)

if ($result -eq 'Yes')
    {
    $ServerSearch = @()
    $ServerSearch = [Microsoft.VisualBasic.Interaction]::InputBox('Quais servidores Exchange receberam a coleta. Exemplo:', 'Exchange Servers', "'SRV-EXCH-01','SRV-EXCH-02'")
    $ServerSearch = $(($ServerSearch -split ",") -replace "'",'')
    Write-Host "yes - Remote Connection"
    $ServerSearch | % {
    $RemoteSession = New-PSSession -ComputerName $ServerSearch
    }
}else{
    $ServerSearch = @()
    $ServerSearch = [Microsoft.VisualBasic.Interaction]::InputBox('Quais servidores Exchange receberam a coleta. Exemplo:', 'Exchange Servers', "'SRV-EXCH-01','SRV-EXCH-02'")
    $ComputerName = "ComputerName"
    $ServerSearch = $("'ComputerName',$ServerSearch")
    $ServerSearch = $(($ServerSearch -split ",") -replace "'",'')
    Write-Host "no - Local Run"
    $RemoteSession = ($ServerSearch | Select-Object -Skip 1 | ConvertFrom-CSV -Header $ComputerName) 
}
#######################################################################################################

$Folder  = [String](Get-Date -Format yyyyMd)
$zipFolder  = [String](Get-Date -Format Md)

$Contador = 0


$RemoteSession | % {
    $Contador++


    if (Get-Item -Path "$($Directory + '\' + $Folder + '\' + $_.ComputerName + '-' + $zipFolder + '.zip')" -ErrorAction SilentlyContinue)
    {
        Write-Host "$Contador - File found ... descompress $($_.ComputerName).zip"
        
        Expand-Archive -path "$($Directory + '\' + $Folder + '\' + $_.ComputerName + '-' + $zipFolder + '.zip')" -destinationpath "$($Directory + '\' + $Folder + '\' + $_.ComputerName + '-' + $zipFolder)"

        $Archives = Get-ChildItem "$($Directory + '\' + $Folder + '\' + $_.ComputerName + '-' + $zipFolder)" | ? {($_.Name -like "*.zip")} 
        
        foreach($Archive in $Archives){
            $FullName = $($Archive.FullName)
            $destinationpath = $($Archive.FullName -replace ".zip",'')

            Expand-Archive -path "$($FullName)" -destinationpath "$($destinationpath)"
        }
    }
    else
    {
        Write-Host "$Contador - Copy and descompress  file server .. $($_.ComputerName)"

        Copy-Item -FromSession $($_) –Path "$($Directory + '\' + $Folder + '\' + $_.ComputerName + '-' + $zipFolder + '.zip')" –Destination "$($Directory + '\' + $Folder + '\')" 
        
        Expand-Archive -path "$($Directory + '\' + $Folder + '\' + $_.ComputerName + '-' + $zipFolder + '.zip')" -destinationpath "$($Directory + '\' + $Folder + '\' + $_.ComputerName + '-' + $zipFolder)"
        
        $Archives = Get-ChildItem "$($Directory + '\' + $Folder + '\' + $_.ComputerName + '-' + $zipFolder)" | ? {($_.Name -like "*.zip")} 
            
            foreach($Archive in $Archives){
            $FullName = $($Archive.FullName)
            $destinationpath = $($Archive.FullName -replace ".zip",'')

            Expand-Archive -path "$($FullName)" -destinationpath "$($destinationpath)"
        }
    }

} 
