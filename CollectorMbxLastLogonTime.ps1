
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Set-ADServerSettings -ViewEntireForest $true


#Servidores
$ServerSearch = 'Name','SRV-EXCH13-01','SRV-EXCH13-02'


$AllMailboxes = Get-Mailbox -ResultSize Unlimited | ? {($ServerSearch -contains $_.ServerName)} | Select-Object DisplayName,Alias,SamAccountName,UserPrincipalName,PrimarySmtpAddress,RecipientTypeDetails,OrganizationalUnit
$Total = $AllMailboxes.count


$Report = [System.Collections.Generic.List[Object]]::new()
$Contador = 0

ForEach ($Mailboxes in $AllMailboxes) {
    $Contador++
    $DisplayName = @($Mailboxes.DisplayName);
    $Alias = @($Mailboxes.Alias);
    $SamAccountName = @($Mailboxes.SamAccountName);
    $UserPrincipalName = @($Mailboxes.UserPrincipalName);
    $PrimarySmtpAddress = @([String]$Mailboxes.PrimarySmtpAddress.Address);
    $RecipientTypeDetails = @($Mailboxes.RecipientTypeDetails);
    $OrganizationalUnit = @($Mailboxes.OrganizationalUnit);
    
    if (!(Get-MailboxStatistics -Identity $([String]$PrimarySmtpAddress)).LastLogonTime)
    {
    Write-Host "login null - [$Contador], aguarde até [$Total]"
        $ReportLine = [PSCustomObject]@{
        DisplayName = $($DisplayName)
        Alias = $($Alias)
        SamAccountName = $($SamAccountName)
        UserPrincipalName = $($UserPrincipalName)
        PrimarySmtpAddress = $($PrimarySmtpAddress)
        RecipientTypeDetails = $($RecipientTypeDetails)
        LastLogonTime = $($Null)
        TotalItemSize = $((Get-MailboxStatistics -Identity $([String]$PrimarySmtpAddress) -ErrorAction SilentlyContinue).TotalItemSize)
        TotalDeletedItemSize = $((Get-MailboxStatistics -Identity $([String]$PrimarySmtpAddress) -ErrorAction SilentlyContinue).TotalDeletedItemSize)
        OrganizationalUnit = $($OrganizationalUnit)
        }
        $Report.Add($ReportLine)
    }
    else
    {
    Write-Host "login value - [$Contador], aguarde até [$Total]"
        $ReportLine = [PSCustomObject]@{
        DisplayName = $($DisplayName)
        Alias = $($Alias)
        SamAccountName = $($SamAccountName)
        UserPrincipalName = $($UserPrincipalName)
        PrimarySmtpAddress = $($PrimarySmtpAddress)
        RecipientTypeDetails = $($RecipientTypeDetails)
        LastLogonTime = $([String](Get-MailboxStatistics -Identity $([String]$PrimarySmtpAddress) -ErrorAction SilentlyContinue).LastLogonTime.ToString('dd/MM/yyyy HH:mm:ss'))
        TotalItemSize = $((Get-MailboxStatistics -Identity $([String]$PrimarySmtpAddress) -ErrorAction SilentlyContinue).TotalItemSize)
        TotalDeletedItemSize = $((Get-MailboxStatistics -Identity $([String]$PrimarySmtpAddress) -ErrorAction SilentlyContinue).TotalDeletedItemSize)
        OrganizationalUnit = $($OrganizationalUnit)
        }
        $Report.Add($ReportLine)
    }

    $DisplayName = $Null;
    $UserPrincipalName = $Null;
    $PrimarySmtpAddress = $Null;
    $RecipientTypeDetails = $Null;
    $OrganizationalUnit = $Null;
}


$Diretorio = "D:\ExchangeLogCollector\CollectorMbx\"
mkdir $Diretorio
$Report | Select-Object * | Export-Csv -Path $($Diretorio + "CollectorMbx_LastLogin.csv") -NoTypeInformation -Encoding UTF8