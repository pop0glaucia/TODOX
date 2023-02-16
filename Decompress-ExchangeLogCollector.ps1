#######################################################################################################
# WinRM - Remote Conections
#######################################################################################################
<#
$ServerSearch = 'SRV-EXCH13-01','SRV-EXCH13-02'
$ServerSearch | % {
    $RemoteSession = New-PSSession -ComputerName $ServerSearch
}
#>
#######################################################################################################
# Local Conections
#######################################################################################################

$ComputerName = "ComputerName"
$ServerSearch = 'ComputerName','SRV-EXCH13-01','SRV-EXCH13-02'
$RemoteSession = ($ServerSearch | Select-Object -Skip 1 | ConvertFrom-CSV -Header $ComputerName)
#######################################################################################################

$Folder  = [String](Get-Date -Format yyyyMd)
$zipFolder  = [String](Get-Date -Format Md)

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
