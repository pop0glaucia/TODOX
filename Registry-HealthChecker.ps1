
Write-Host 'NETFramework'
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319' -Name 'SchUseStrongCrypto' -Type DWord -Value '1'
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319' -Name 'SchUseStrongCrypto' -Type DWord -Value '1'
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319' -Name 'SystemDefaultTlsVersions' -type DWord -Value '1'
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319' -Name 'SystemDefaultTlsVersions' -type DWord -Value '1'

Write-Host 'KeepAliveTime'
Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\TcpIp\Parameters' -Name 'KeepAliveTime' -Type DWord -Value '1800000'

Write-Host 'PnPCapabilities'
Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}' -Name 'PnPCapabilities' -Type DWord -Value '280'

Write-Host 'Disabled - TLS 1.2'
Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server' -Name 'Enabled' -Type DWord -Value '0'
Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server' -Name 'DisabledByDefault' -Type DWord -Value '1'
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client' -Force
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client' –PropertyType 'DWORD' -Name 'Enabled' -Value '0'
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client' -PropertyType 'DWORD' -Name 'DisabledByDefault' -Value '1'
	
Write-Host 'Disabled - TLS 1.2'		
Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server' -Name 'Enabled' -Type DWord -Value '0'
Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server' -Name 'DisabledByDefault' -Type DWord -Value '1'
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client' -Force
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client' –PropertyType 'DWORD' -Name 'Enabled' -Value '0'
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client' -PropertyType 'DWORD' -Name 'DisabledByDefault' -Value '1'

Write-Host 'Enabled - TLS 1.2'
New-Item 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -Force
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -PropertyType 'DWORD' -Name 'Enabled' -Value '1'
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' –PropertyType 'DWORD' -Name 'DisabledByDefault' -Value '0'

Write-Host 'Disabled Ipv6'
Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters' -Name 'DisabledComponents' -Type DWord -Value '255'

Write-Host 'Disabled SMB1 Protocol'
Disable-WindowsOptionalFeature -Online -FeatureName smb1protocol
Set-SmbServerConfiguration -EnableSMB1Protocol $false