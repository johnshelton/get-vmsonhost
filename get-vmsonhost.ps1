<#
=======================================================================================
File Name: get-vmsonhost.ps1
Created on: 
Created with VSCode
Version 1.0
Last Updated: 
Last Updated by: John Shelton | c: 260-410-1200 | e: john.shelton@lucky13solutions.com

Purpose: Collect list of all VMs running by VMWare Host in VCenter

Notes: 

Change Log:


=======================================================================================
#>
#
# Define Parameter(s)
#
param (
  [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
  [string[]] $VCenterServers = $(throw "-VCenterServers is required.  Pass as array.")
)
#
Clear-Host
# Define Output Variable
$AllVMInfo = @()
#
# Load VMWare PSSnapin
#
Add-PSSnapin VMWare.VimAutomation.Core
#
# Define Output Variables
#
$ExecutionStamp = Get-Date -Format yyyyMMdd_hh-mm-ss
$path = "c:\temp\"
$FilenamePrepend = 'temp_'
$FullFilename = "get-vmsonhost.ps1"
$FileName = $FullFilename.Substring(0, $FullFilename.LastIndexOf('.'))
$FileExt = '.xlsx'
$OutputFile = $path + $FilenamePrePend + '_' + $FileName + '_' + $ExecutionStamp + $FileExt
$PathExists = Test-Path $path
IF($PathExists -eq $False)
  {
  New-Item -Path $path -ItemType  Directory
  }
#
$CountVCenterServers = $VCenterServers.Count
foreach($VCenterServer in $VCenterServers){
  connect-viserver $VCenterServer
  $VCenterServersProcessed++
  $PercentVCenterServers = ($VCenterServersProcessed/$CountVCenterServers*100)
  Write-Progress -Activity "Processing through VCenter Servers" -PercentComplete $PercentVCenterServers -CurrentOperation "Processing $VCenterServer" -ID 1
  $VMHosts = Get-VMHost | Where-Object {$_.ConnectionState -eq "Connected"}
  $VMs = Get-VM
  $CountVMs = $VMs.Count
  ForEach ($VM in $VMs){
    $VMsProcessed++
    $PercenetVMsProcessed = ($VMsProcessed/$CountVMs*100)
    Write-Progress -Activity "Processing through all VMs on $VCenterServer" -PercentComplete $PercenetVMsProcessed -CurrentOperation "Processing $VM" -ParentId 1
    $VMDNS = Resolve-DnsName $VM.Name -ErrorAction SilentlyContinue
    IF(!$VMDNS.Name) {$VMConnected = "No DNS Name Found"}
    Else {$VMConnected = Test-Connection $VMDNS.Name -Count 1 -Quiet}
    $results = New-Object psobject
    $results | Add-Member -MemberType NoteProperty -Name "Name" -Value $VM.Name
    $results | Add-Member -MemberType NoteProperty -Name "VMHost" -Value $VM.VMHost
    $results | Add-Member -MemberType NoteProperty -Name "DNS Name" -Value $VMDNS.Name
    $results | Add-Member -MemberType NoteProperty -Name "IP" -Value $VMDNS.IPAddress
    $results | Add-Member -MemberType NoteProperty -Name "RepliedToPing" -Value $VMConnected
    $results | Add-Member -MemberType NoteProperty -Name "VCenterServer" -Value $VM.VMHost
    $results | Add-Member -MemberType NoteProperty -Name "VMWareFolder" -Value $VM.Folder
    $results | Add-Member -MemberType NoteProperty -Name "VMPowerState" -Value $VM.PowerState
    $results | Add-Member -MemberType NoteProperty -Name "VMGuestInfo" -Value $VM.Guest
    $AllVMInfo += $results
  }
  $AllVMInfo | Export-Excel -Path $OutputFile -WorkSheetname $VCenterServer -TableName $VCenterServer -TableStyle Medium4
  $VMsProcessed = 0
}