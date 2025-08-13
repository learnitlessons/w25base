New-VHD -ParentPath "D:\VM\vhd\w25gui.vhdx" -Path D:\vm\w25-lit-svr1\w25-lit-svr1.vhdx -Differencing -Verbose
New-VM -Name w25-lit-svr1 -MemoryStartupBytes 4Gb -VHDPath D:\vm\w25-lit-svr1\w25-lit-svr1.vhdx -Path D:\vm\w25-lit-svr1  -Generation 2 -SwitchName Ext

set-vm w25-lit-svr1 -CheckpointType Disabled
start-vm w25-lit-svr1

# Enter-PSSession
$Credentials=Get-Credential
etsn -VMName w25-lit-svr1 -Credential $Credentials

Rename-Computer -NewName "w25-lit-svr1" -Restart -Force

# Configure IP settings
$interfaceIndex = (Get-NetAdapter | Where-Object {$_.Status -eq "Up"}).InterfaceIndex
New-NetIPAddress -InterfaceIndex $interfaceIndex -IPAddress 192.168.1.23 -PrefixLength 24 -DefaultGateway 192.168.1.1
Set-DnsClientServerAddress -InterfaceIndex $interfaceIndex -ServerAddresses 192.168.1.22

# installing active directory domain services role
install-windowsfeature -name ad-domain-services -includemanagementtools
import-module addsdeployment
install-addsforest -creatednsdelegation:$false -databasepath "c:\windows\ntds" -domainmode "Win2025" -domainname "learnitlessons.com" -domainnetbiosname "lit" -forestmode "Win2025" -installdns:$true -logpath "c:\windows\ntds" -norebootoncompletion:$false -sysvolpath "c:\windows\sysvol" -force:$true

# adding member server to the learnitlessons.com domain
add-computer  -domainname learnitlessons.com -credential lit\administrator -verbose -restart -forc

Stop-VM -Name w25-lit-svr1 -Force
Set-VM w25-lit-svr1 -CheckpointType Production
Get-VM -Name w25-lit-svr1
Checkpoint-VM -Name w25-lit-svr1 -SnapshotName 'initial'
Restore-VMSnapshot -VMName w25-lit-svr1 -Name "initial" -Confirm:$false
Start-VM -Name w25-lit-svr1

Stop-VM -Name w25-lit-svr1 
Remove-VM w25-lit-svr1 -Force
Remove-Item -Recurse D:\vm\w25-lit-svr1 -Force


$Credentials=Get-Credential
etsn -VMName lit-svr1 -Credential $Credentials

get-vm LIT-* | Checkpoint-VM -SnapshotName 'temp'
get-vm LIT-* | Restore-VMSnapshot -Name "Lab ..." -Confirm:$false
get-vm LIT-* | Start-VM


Install-WindowsFeature -Name DNS -IncludeManagementTools
Add-DnsServerSecondaryZone -Name "learnitlessons.com" -ZoneFile "learnitlessons.com.dns" -MasterServers 192.168.2.22
