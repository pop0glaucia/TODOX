#######################################################################################################
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$result = [System.Windows.Forms.MessageBox]::Show('O Script ExchangeLogCollector foi executado a partir de uma maquina remota?' , "Tipo de Coleta" , 4)

if ($result -eq 'Yes')
    {
    $ServerSearch = @()
    $ServerSearch = [Microsoft.VisualBasic.Interaction]::InputBox('Quais servidores Exchange receberam a coleta. Exemplo:', 'Exchange Servers', "'SRV-EXCH13-01','SRV-EXCH13-02'")
    Write-Host "yes - Remote Connection"
    $ServerSearch | % {
    $RemoteSession = New-PSSession -ComputerName $ServerSearch
    }
}else{
    $ServerSearch = @()
    $ServerSearch = [Microsoft.VisualBasic.Interaction]::InputBox('Quais servidores Exchange receberam a coleta. Exemplo:', 'Exchange Servers', "'SRV-EXCH13-01','SRV-EXCH13-02'")
    $ComputerName = "ComputerName"
    $ServerSearch = $("'ComputerName',$ServerSearch")
    $ServerSearch = $(($ServerSearch -split ",") -replace "'",'')
    Write-Host "no - Local Run"
    $RemoteSession = ($ServerSearch | Select-Object -Skip 1 | ConvertFrom-CSV -Header $ComputerName) 
}
#######################################################################################################

$Folder  = [String](Get-Date -Format yyyyMd)

$Contador = 0

$RemoteSession | % {
    $Contador++

    Copy-Item -FromSession $($_) –Path "$('D:\ExchangeLogCollector\' + $Folder + '\' + $_.ComputerName + '-' + $zipFolder + '.zip')" –Destination "$('D:\ExchangeLogCollector\' + $Folder + '\')" 
    Expand-Archive -path "$('D:\ExchangeLogCollector\' + $Folder + '\' + $_.ComputerName + '-' + $zipFolder + '.zip')" -destinationpath "$('D:\ExchangeLogCollector\' + $Folder + '\' + $_.ComputerName + '-' + $zipFolder)"
    $Archives = Get-ChildItem "$('D:\ExchangeLogCollector\' + $Folder + '\' + $_.ComputerName + '-' + $zipFolder)" | ? {($_.Name -like "*.zip")}
    
    foreach($Archive in $Archives){
        $FullName = $($Archive.FullName)
        $Directory = $($Archive.FullName -replace ".zip",'')

        Expand-Archive -path "$($FullName)" -destinationpath "$($Directory)"
    }
    Write-Host "$Contador - Copy file server .. $($_.ComputerName)"
}
