
Write-Host "############################################"
Write-Host "Run this script in mode administrative"
Write-Host "############################################"


$ServerSearch = 'SRV-EXCH13-01','SRV-EXCH13-02'

$Folder  = [String](Get-Date -Format yyyyMd)
$zipFolder  = [String](Get-Date -Format Md)

$ProtocolPath = @()

$ServerSearch | %{

    $ProtocolPath += Get-ChildItem "$('D:\ExchangeLogCollector\' + $Folder + '\' + $($_) + '-' + $zipFolder)" | ? {($_.Name -notlike "*.zip") -and (($_.Name -notlike 'Windows_Event_Logs') -and ($_.Name -notlike "*.txt") -and ($_.Name -notlike "*.xml"))}
}


$logs = Get-ChildItem ($ProtocolPath).FullName *.log

$logs | % {
    $FolderSanitized = [String]@($_.Directory -replace $Folder,'Sanitized')
    $Export = [String]@($FolderSanitized.Split("\")[0..2] + $FolderSanitized.Split("\")[-1] + $FolderSanitized.Split("\")[-2] -join "\")
    $Export = [String]@($Export -replace "$('-' + $zipFolder)",'')
    $Name = [String]@($_.BaseName)
    $Path = [String]($Export + '\' + $Name + '.csv')

    New-Item -ItemType Directory -Path "$($Export)" -ErrorAction SilentlyContinue
    
    if ($Export -like "*HTTPERR_Logs*")
    {
        Get-Content $_.FullName | % {
            $IsHeaderParsed = $false

            if ($_ -like '#Fields: *' -and !$IsHeaderParsed){
                $_ -replace '^#Fields: '
                $IsHeaderParsed = $true
            }
            else{
                $_
            }

        } | ConvertFrom-Csv -Delimiter " " | Export-Csv -Path "$($Path)" -NoTypeInformation -Encoding UTF8
    }
    elseif ($Export -like "*IIS_W3SVC*")
    {
        Get-Content $_.FullName | % {
            $IsHeaderParsed = $false

            if ($_ -like '#Fields: *' -and !$IsHeaderParsed){
                $_ -replace '^#Fields: '
                $IsHeaderParsed = $true
            }
            else{
                $_
            }

        } | ConvertFrom-Csv -Delimiter " " | Export-Csv -Path "$($Path)" -NoTypeInformation -Encoding UTF8
    }
    else
    {
        Get-Content $_.FullName | % {
            $IsHeaderParsed = $false

            if ($_ -like '#Fields: *' -and !$IsHeaderParsed){
                $_ -replace '^#Fields: '
                $IsHeaderParsed = $true
            }
            else{
                $_
            }

        } | ConvertFrom-Csv | Export-Csv -Path "$($Path)" -NoTypeInformation -Encoding UTF8
    }

}
$Directory = @()
$Directory = Get-ChildItem -Path 'D:\ExchangeLogCollector\Sanitized\' | ? {(($_.Name -like "*W3SVC*") -or ($_.Name -like "*HTTPERR*"))} | Select-Object FullName,Name

$Directory | % {
    $Path = [String]@($_.FullName)
    $Name = [String]@($_.Name)
    
    if ($Name -like "*W3SVC*")
    {
    New-Item -ItemType Directory -Path 'D:\ExchangeLogCollector\Sanitized\IISLogs\' -ErrorAction SilentlyContinue
    Move-Item -Path $Path -Destination 'D:\ExchangeLogCollector\Sanitized\IISLogs\'
    Write-Host "Move Directory $($Name) to 'D:\ExchangeLogCollector\Sanitized\IISLogs\"
    }
    else
    {
    New-Item -ItemType Directory -Path 'D:\ExchangeLogCollector\Sanitized\RpcLogs\' -ErrorAction SilentlyContinue
    Move-Item -Path $Path -Destination 'D:\ExchangeLogCollector\Sanitized\RpcLogs\'
    Write-Host "Move Directory $($Name) to 'D:\ExchangeLogCollector\Sanitized\RpcLogs\"
    }
}

<#
$Directory = $logs | Select-Object -Unique Directory
$Directory = $Directory | Select-Object @{Name = 'Directory'; Expression = {($_.Directory -replace $Folder,'Sanitized')}}
$Directory = $Directory | Select-Object @{Name = 'Directory'; Expression = {(($_.Directory).Split("\")[0..3] -join "\")}}
$Directory = $Directory | Select-Object -Unique Directory


$Directory | % {
    $Path = [String]@($_.Directory)
    $Name = [String]@($($Path.Split("\")[-1] -replace "$('-' + $zipFolder)",''))
    
    Write-Host "Rename folder ... $($Path.Split("\")[-1]) to --> $Name"

    Rename-Item "$($Path)" -NewName "$($Name)"

}
#>

Write-Host "############################################"
Write-Host "Finish"
Write-Host "############################################"