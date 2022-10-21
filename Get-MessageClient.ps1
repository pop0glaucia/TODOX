#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Write-Host "Criar variavel PATH"
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$Diretorio = [Microsoft.VisualBasic.Interaction]::InputBox('Informe o nome do diretório onde deseja armazenar os dados. Exemplo:', 'Diretorio ou Path', "C:\Scripts\")
cd $Diretorio

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Write-Host "Criar variavel contendo a lista de ultimo login das caixas de correio"
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$File = [Microsoft.VisualBasic.Interaction]::InputBox('Informe o nome do arquivo que contem os dados de [MbxLastLogin.csv]. Exemplo:', 'Arquivo .CSV', "21hr28min_20_10_2022_CollectorMbx_LastLogin")

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Write-Host "Criar variavel com o nome do servidor"
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$Server = [Microsoft.VisualBasic.Interaction]::InputBox('SERVIDOR EXCHANGE A SER COLEADO:', 'EXCHANGE SERVER', "$($env:COMPUTERNAME)")

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Write-Host "Criar variavel com a quantidade de dias de logs a serem coletados"
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$Date = [Microsoft.VisualBasic.Interaction]::InputBox('Informe a quantidade minima de dias do log ser coletado. Exemplo:', 'Dia(s) de coleta', "2")

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Write-Host "Importar Modulos da Pshell do Exchange 2013-2016-2019"
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Set-ADServerSettings -ViewEntireForest $true

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Write-Host "Criar função de coleta para extração dos dados de MessageTrackingLog"
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

function Get-MessageClientType
    {
        $MessageTrackingLog = @($input) | ? {$_.SourceContext -match "ClientType"}
        $Output = @()
        foreach ($Message in $MessageTrackingLog)
            {
                $ClientType = $Message.SourceContext -split “,” | ? {$_ -match "ClientType"}
                $ClientType = $ClientType -replace (" ClientType:","") 
            
                $ResultClientType = @()
                $ClientType | ? {
                    if ($_ -eq 'MOMT'){$ResultClientType += 'OUTLOOK DESKTOP'}
                    elseif($_ -like 'AirSync'){$ResultClientType += 'OUTLOOK MOBILE'}
                    else{$ResultClientType += "$($_)"}
                }
                #>

                $ResultClientType = ($ResultClientType -split "{")[0] -replace "}",''

                             
                $OutputLine = New-Object System.Object
                $OutputLine | Add-Member -Type NoteProperty -Name TimeStamp -Value $Message.TimeStamp
                $OutputLine | Add-Member -Type NoteProperty -Name Sender -Value $Message.Sender
                $OutputLine | Add-Member -Type NoteProperty -Name Recipients -Value $Message.Recipients
                $OutputLine | Add-Member -Type NoteProperty -Name MessageSubject -Value $Message.MessageSubject
                $OutputLine | Add-Member -Type NoteProperty -Name EventId -Value $Message.EventId
                $OutputLine | Add-Member -Type NoteProperty -Name ServerIp -Value $Message.ServerIp
                $OutputLine | Add-Member -Type NoteProperty -Name ClientIp -Value $Message.ClientIp
                $OutputLine | Add-Member -Type NoteProperty -Name ClientType -Value $ResultClientType
                $Output += $OutputLine
            }
        $Output
    }

function Get-MessageClientProtocol
    {
        $MessageTrackingLog = @($input) | ? {$_.SourceContext -match "ClientSubmitTime"} | ? {($_.EventId -eq "SEND")} | ? {(($_.SourceContext -notlike $null) -and ($_.MessageInfo -notlike $null))} | ? {($_.MessageInfo -notlike "*MTS*")}
        $Output = @()
        foreach ($Message in $MessageTrackingLog)
            {
                $ClientProtocol = ($Message.SourceContext  -split “;”)[2] -replace "ClientSubmitTime:",'IMAP, POP or SMTP'
                $OutputLine = New-Object System.Object
                $OutputLine | Add-Member -Type NoteProperty -Name TimeStamp -Value $Message.TimeStamp
                $OutputLine | Add-Member -Type NoteProperty -Name Sender -Value $Message.Sender
                $OutputLine | Add-Member -Type NoteProperty -Name Recipients -Value $Message.Recipients
                $OutputLine | Add-Member -Type NoteProperty -Name MessageSubject -Value $Message.MessageSubject
                $OutputLine | Add-Member -Type NoteProperty -Name EventId -Value $Message.EventId
                $OutputLine | Add-Member -Type NoteProperty -Name ServerIp -Value $Message.ServerIp
                $OutputLine | Add-Member -Type NoteProperty -Name ClientIp -Value $Message.ClientIp
                $OutputLine | Add-Member -Type NoteProperty -Name ClientType -Value $ClientProtocol
                $Output += $OutputLine
            }
        $Output
    }


#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Write-Host "Testar dados do coletor MessageTrackingLog"
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

$MbxTeste = Get-Mailbox -ResultSize 12 | % {
        $UserPrincipalName = "$($_.UserPrincipalName)";
        $PrimarySmtpAddress = "$($_.PrimarySmtpAddress)";
        $RecipientTypeDetails = "$($_.RecipientTypeDetails)";
        $OrganizationalUnit = "$($_.OrganizationalUnit)";
        $PrimarySmtpAddress | Get-MailboxStatistics | Select-Object  `
        DisplayName, 
        @{N='UserPrincipalName'; E={$UserPrincipalName}}, 
        @{N='PrimarySmtpAddress'; E={$PrimarySmtpAddress}}, 
        @{N='LastLogonTime'; E={$($_.LastLogonTime)}}, 
        @{N='RecipientTypeDetails'; E={$RecipientTypeDetails}}, 
        TotalItemSize, 
        TotalDeletedItemSize, 
        @{N='OrganizationalUnit'; E={$OrganizationalUnit}} 
        ;} | ?  {($_.'LastLogonTime' -gt (get-date).AddDays(-[int]$Date))}
        


# Quantidade de Mailboxes a receber o Teste
$Total = $MbxTeste.count
Write-Host "Teste - Total de caixas de correio para teste do script $([int]$Total)"

# Pesquisar Data
$SearchDate = $([string]((get-date).AddDays(-[int]$Date)).ToString("MM/dd/yyyy hh:mm:ss tt"))
Write-Host "Teste - Pesquisando data $SearchDate"

# Inciando Coleta
Write-Host "Teste - Inicando Coleta .. Aguarde"

# Zerando Contador
$Contador = 0
Write-Host "Teste - Zerando contador"

foreach ($Teste in $MbxTeste) {
    $Contador++
    $PrimarySmtpAddress = @($Teste.PrimarySmtpAddress);
    $LastLogonTime = @($Teste.LastLogonTime);

    if (!(Get-MessageTrackingLog -Server $Server -Start (get-date).AddDays(-[int]$Date) -Sender "$($PrimarySmtpAddress)" -ResultSize Unlimited | Select-Object `
            @{N='TimeStamp'; E={$_.TimeStamp}} ,      
            @{N='Sender'; E={$_.Sender}},            
            @{N='Recipients'; E={$_.Recipients}},   
            @{N='MessageSubject'; E={$_.MessageSubject}},
            @{N='EventId'; E={$_.EventId}},
            @{N='ServerIp'; E={$_.ServerIp}},
            @{N='ClientIp'; E={$_.ClientIp}},             
            @{N='ClientType'; E={$_.ClientType}},
            @{N='SourceContext'; E={$_.SourceContext}} | Get-MessageClientType ))
    {  
        $PrimarySmtpAddress | Select-Object `
            @{N='ProgressColeta'; E={$("$([String]$Contador++)" + '..de..' + "$Total")}},
            @{N='TimeStamp'; E={''}} ,      
            @{N='Sender'; E={$PrimarySmtpAddress}},            
            @{N='Recipients'; E={''}},   
            @{N='MessageSubject'; E={''}},
            @{N='EventId'; E={''}},
            @{N='ServerIp'; E={''}},
            @{N='ClientIp'; E={''}},             
            @{N='ClientType'; E={''}},
            @{N='Note'; E={"A caixa NAO enviou mensagens utilizando uma conexao HTTPS no(s) ultimo(s) $([int]$Date) dias(s)."}},
            @{N='MailboxLogon'; E={$LastLogonTime}} | ft -AutoSize ProgressColeta,TimeStamp,Sender,Recipients,MessageSubject,ClientType,Note,MailboxLogon
    }
    else
    {
        Get-MessageTrackingLog -Server $Server -Start (get-date).AddDays(-[int]$Date) -Sender "$($PrimarySmtpAddress)" -ResultSize Unlimited | Select-Object `
            @{N='TimeStamp'; E={$_.TimeStamp}} ,      
            @{N='Sender'; E={$_.Sender}},            
            @{N='Recipients'; E={$_.Recipients}},   
            @{N='MessageSubject'; E={$_.MessageSubject}},
            @{N='EventId'; E={$_.EventId}},
            @{N='ServerIp'; E={$_.ServerIp}},
            @{N='ClientIp'; E={$_.ClientIp}},             
            @{N='ClientType'; E={$_.ClientType}},
            @{N='SourceContext'; E={$_.SourceContext}} | Get-MessageClientType  | ? {($_.EventId -notlike 'SUBMITDEFER')}  | Select-Object @{N='ProgressColeta'; E={$("$([int]$Contador++)" + '..de..' + "$Total")}},*,@{N='Note'; E={'A caixa ENVIOU mensagens utilizando uma conexao HTTPS'}},@{N='MailboxLogon'; E={$($LastLogonTime)}} | ft -AutoSize ProgressColeta,TimeStamp,Sender,Recipients,MessageSubject,ClientType,Note,MailboxLogon
    }

    if (!(Get-MessageTrackingLog -Server $Server -Start (get-date).AddDays(-[int]$Date) -Sender "$($PrimarySmtpAddress)" -ResultSize Unlimited | Select-Object *  | Get-MessageClientProtocol ))
    {
        $PrimarySmtpAddress | Select-Object `
            @{N='ProgressColeta'; E={$("$([String]$Contador++)" + '..de..' + "$Total")}},
            @{N='TimeStamp'; E={''}} ,      
            @{N='Sender'; E={$PrimarySmtpAddress}},            
            @{N='Recipients'; E={''}},   
            @{N='MessageSubject'; E={''}},
            @{N='EventId'; E={''}},
            @{N='ServerIp'; E={''}},
            @{N='ClientIp'; E={''}},             
            @{N='ClientType'; E={''}},
            @{N='Note'; E={"A caixa NAO enviou mensagens utilizando uma conexao IMAP, POP or SMTP no(s) ultimo(s) $([int]$Date) dias(s)."}},
            @{N='MailboxLogon'; E={$LastLogonTime}} | ft -AutoSize ProgressColeta,TimeStamp,Sender,Recipients,MessageSubject,ClientType,Note,MailboxLogon
    }
    else
    {
        Get-MessageTrackingLog -Server $Server -Start (get-date).AddDays(-[int]$Date) -Sender "$($PrimarySmtpAddress)" -ResultSize Unlimited | Select-Object * | Get-MessageClientProtocol  | Select-Object @{N='ProgressColeta'; E={$("$([int]$Contador++)" + '..de..' + "$Total")}},*,@{N='Note'; E={'A caixa ENVIOU mensagens utilizando uma conexao IMAP, POP or SMTP'}},@{N='MailboxLogon'; E={$($LastLogonTime)}} | ft -AutoSize ProgressColeta,TimeStamp,Sender,Recipients,MessageSubject,ClientType,Note,MailboxLogon
    }
}


# Coleta Finalizada
Write-Host "Teste - Coleta .. Finalizada"

#Get-ExchangeServer | Get-MessageTrackingLog -Start "10/20/2022 0:01:00 AM" -Sender joao.santos@CONTOSO.COM -ResultSize Unlimited  | Get-MessageClientType  | ? {($_.EventId -notlike 'SUBMITDEFER')}  | ft -AutoSize
#Get-ExchangeServer | Get-MessageTrackingLog -Start "10/20/2022 0:01:00 AM" -Sender NO-REPLY@CONTOSO.COM -ResultSize Unlimited  | Get-MessageClientType  | ? {($_.EventId -notlike 'SUBMITDEFER')}  | ft -AutoSize
#Get-ExchangeServer | Get-MessageTrackingLog -Start "10/20/2022 0:01:00 AM" -Sender NO-REPLY@CONTOSO.COM -ResultSize Unlimited  | Get-MessageClientProtocol | ft -AutoSize

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Ler: https://learn.microsoft.com/en-us/previous-versions/tn-archive/cc539064(v=technet.10)?redirectedfrom=MSDN
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Carregar Lista
$Lista = Import-Csv -Path $($Diretorio + $File + '.csv')

# Inciando Coleta Full - Criar variavel com o total de caixas de correio hospedadas no Exchange Server
$Total = $Lista.count
Write-Host "Coleta Full - total de caixas de correio hospedadas no Exchange Server [$([int]$Total)]."


# Zerando Contador
$Contador = 0
$ContadorFull = 1
Write-Host "Coleta Full - Zerando contador"

# Limpando Lista para exportacao
$Report = @()

# Iniciando Coleta Fullt
Write-Host "Coleta Full .. Iniciado"

foreach ($Mailbox in $Lista) {
    $Contador++
    $PrimarySmtpAddress = @($Mailbox.PrimarySmtpAddress);
    $LastLogonTime = @($Mailbox.LastLogonTime);
    
    Write-Host "Total de Caixas de Correio $([int]$Total) - $('Inicando coleta .. -> '+ "$([String]$ContadorFull++)" + '..de..' + "$Total")"

    if (!(Get-MessageTrackingLog -Server $Server -Start (get-date).AddDays(-[int]$Date) -Sender "$($PrimarySmtpAddress)" -ResultSize Unlimited | Select-Object `
            @{N='TimeStamp'; E={$_.TimeStamp}} ,      
            @{N='Sender'; E={$_.Sender}},            
            @{N='Recipients'; E={$_.Recipients}},   
            @{N='MessageSubject'; E={$_.MessageSubject}},
            @{N='EventId'; E={$_.EventId}},
            @{N='ServerIp'; E={$_.ServerIp}},
            @{N='ClientIp'; E={$_.ClientIp}},             
            @{N='ClientType'; E={$_.ClientType}},
            @{N='SourceContext'; E={$_.SourceContext}} | Get-MessageClientType ))
    {  
        $Arr1 = $PrimarySmtpAddress | Select-Object `
            @{N='ProgressColeta'; E={$("$([String]$Contador++)" + '..de..' + "$Total")}},
            @{N='TimeStamp'; E={''}} ,      
            @{N='Sender'; E={$PrimarySmtpAddress}},            
            @{N='Recipients'; E={''}},   
            @{N='MessageSubject'; E={''}},
            @{N='EventId'; E={''}},
            @{N='ServerIp'; E={''}},
            @{N='ClientIp'; E={''}},             
            @{N='ClientType'; E={''}},
            @{N='Note'; E={"A caixa NAO enviou mensagens utilizando uma conexao HTTPS no(s) ultimo(s) $([int]$Date) dias(s)."}},
            @{N='MailboxLogon'; E={$LastLogonTime}}
        $Report += $Arr1
    }
    else
    {
        $Arr2 = Get-MessageTrackingLog -Server $Server -Start (get-date).AddDays(-[int]$Date) -Sender "$($PrimarySmtpAddress)" -ResultSize Unlimited | Select-Object `
            @{N='TimeStamp'; E={$_.TimeStamp}} ,      
            @{N='Sender'; E={$_.Sender}},            
            @{N='Recipients'; E={$_.Recipients}},   
            @{N='MessageSubject'; E={$_.MessageSubject}},
            @{N='EventId'; E={$_.EventId}},
            @{N='ServerIp'; E={$_.ServerIp}},
            @{N='ClientIp'; E={$_.ClientIp}},             
            @{N='ClientType'; E={$_.ClientType}},
            @{N='SourceContext'; E={$_.SourceContext}} `
            | Get-MessageClientType  | ? {($_.EventId -notlike 'SUBMITDEFER')}  | Select-Object `
            @{N='ProgressColeta'; E={$("$([int]$Contador++)" + '..de..' + "$Total")}},
            *,
            @{N='Note'; E={'A caixa ENVIOU mensagens utilizando uma conexao HTTPS'}},
            @{N='MailboxLogon'; E={$($LastLogonTime)}}
        $Report += $Arr2
    }

    if (!(Get-MessageTrackingLog -Server $Server -Start (get-date).AddDays(-[int]$Date) -Sender "$($PrimarySmtpAddress)" -ResultSize Unlimited | Select-Object *  | Get-MessageClientProtocol ))
    {
        $Arr3 = $PrimarySmtpAddress | Select-Object `
            @{N='ProgressColeta'; E={$("$([String]$Contador++)" + '..de..' + "$Total")}},
            @{N='TimeStamp'; E={''}} ,      
            @{N='Sender'; E={$PrimarySmtpAddress}},            
            @{N='Recipients'; E={''}},   
            @{N='MessageSubject'; E={''}},
            @{N='EventId'; E={''}},
            @{N='ServerIp'; E={''}},
            @{N='ClientIp'; E={''}},             
            @{N='ClientType'; E={''}},
            @{N='Note'; E={"A caixa NAO enviou mensagens utilizando uma conexao IMAP, POP or SMTP no(s) ultimo(s) $([int]$Date) dias(s)."}},
            @{N='MailboxLogon'; E={$LastLogonTime}}
        $Report += $Arr3
    }
    else
    {
        $Arr4 = Get-MessageTrackingLog -Server $Server -Start (get-date).AddDays(-[int]$Date) -Sender "$($PrimarySmtpAddress)" -ResultSize Unlimited | Select-Object *   | Get-MessageClientProtocol  | Select-Object `
        @{N='ProgressColeta'; E={$("$([int]$Contador++)" + '..de..' + "$Total")}},
        *,
        @{N='Note'; E={'A caixa ENVIOU mensagens utilizando uma conexao IMAP, POP or SMTP'}},
        @{N='MailboxLogon'; E={$($LastLogonTime)}}
        $Report += $Arr4
    }
}


# Coleta Fullt - Total de itens encontrados
Write-Host "Coleta Fullt - Total de itens encontrados [$([int]$Report.count)]."


#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Write-Host "Coleta Full - Criar variavel PATH para Exporte dos dados coletado"
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$exDiretorio = [Microsoft.VisualBasic.Interaction]::InputBox('Informe o nome do diretório onde deseja armazenar os dados exportados. Exemplo:', 'Diretorio ou Path', "C:\Scripts\")
cd $exDiretorio

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Write-Host "Coleta Full - Criar variavel contendo o nome do arquivo a ser exportado"
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$exFile = [Microsoft.VisualBasic.Interaction]::InputBox('Informe o nome do arquivo a ser exportado. Exemplo:', 'Arquivo .CSV', "CollectorMbx_ClientTypeSendMail")

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Exportando dados
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Write-Host "Coleta Full - Exportando dados"

# Data da Coleta
$dt=get-date -Format "HH'hr'mm'min'_dd_MM_yyyy_"
Write-Host "Coleta Full - Data da Coleta $($dt)"

# Iniciando Export dos dados coletados
Write-Host "Coleta Full - Iniciando Export dos dados coletados"

$Report | Select-Object `
    @{Name = 'TimeStamp'; Expression = {$_.TimeStamp}}, `
    @{Name = 'Sender'; Expression = {$_.Sender}}, `
    @{Name = 'Recipients'; Expression = {$_.Recipients}}, `
    @{Name = 'MessageSubject'; Expression = {$_.MessageSubject}}, `
    @{Name = 'EventId'; Expression = {$_.EventId}}, `
    @{Name = 'ServerIp'; Expression = {$_.ServerIp}}, `
    @{Name = 'ClientIp'; Expression = {$_.ClientIp}}, `
    @{Name = 'ClientType'; Expression = {$_.ClientType}}, `
    @{Name ='Note'; Expression={$_.Note}},
    @{N='MailboxLogon'; E={$_.MailboxLogon}} `
    | Export-Csv -Path $($exDiretorio + $dt + $exFile + '.csv') -NoTypeInformation -Encoding UTF8 

# Export FINALIZADO
Write-Host "Coleta Full - Export FINALIZADO"



#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- 
