<#
.Synopsis
   Creates a Hyper-V Exchange Lab Environment with customizable number of Exchange Servers
.DESCRIPTION
   Creates a Hyper-V Exchange Lab Environment with customizable number of Exchange Servers
.EXAMPLE
   ./New-ExchangeLabEnvironment.ps1
#>
param(
    [Parameter(Mandatory = $True)]  # How many Exchange servers to create
    [int]$NumberOfExchangeServers,

    [Parameter(Mandatory = $True)]  # What kind of Virtual Hard Drive to create
    [ValidateSet("DynamicWithoutSource", "Differencing", "DynamicWithSource", "FixedWithoutSource", "FixedWithSource")]
    [string]$VirtualHardDriveType,

    [Parameter(Mandatory = $False)] # Where is the hard drive to use as a base located. Only needed if $VirtualHardDriveType is set to "Differencing"
    [string]$DifferencingParentPath,

    [Parameter(Mandatory = $False)] # Where is the hard drive to use as a source located. Only needed if $VirtualHardDriveType is set to "DynamicWithSource" or "FixedWithSource"
    [string]$SourceDisk,

    [Parameter(Mandatory = $False)] # System drive size
    [long]$SizeSystemDrive = 100GB,

    [Parameter(Mandatory = $False)] # Data drive size
    [long]$SizeDataDrive = 100GB,

    [Parameter(Mandatory = $False)] # RAM size
    [long]$SizeRam = 4GB,

    [Parameter(Mandatory = $False)] # Virtual Machine name; Incremental number will be added
    [string]$VirtualMachineName = "EX",

    [Parameter(Mandatory = $False)] # Which ethernet card to use
    [string]$VirtualMachineNetworkInterfaceName = "LAB",

    [Parameter(Mandatory = $False)] # Where to store log file
    [string]$LogFolder = "$($env:windir)\Logs\Software"
)

$ErrorActionPreference = "Stop"

Function Write-Log
{
    param(
        [Parameter(Mandatory = $True)]
        [string]$LogPath,

        [Parameter(Mandatory = $True)]
        [string]$Message,

        [Parameter(Mandatory = $True)]
        [ValidateRange(1,3)]
        [int]$Severity,

        [Parameter(Mandatory = $True)]
        [string]$Program,

        [Parameter(Mandatory = $True)]
        [string]$LogFileComponent = "Unknown",

        [Parameter()]
        [bool]$Append = $True
    )

    $TimeZoneBias = Get-WmiObject -Query "Select Bias from Win32_TimeZone" -ComputerName $env:COMPUTERNAME
    $Date1 = Get-Date -Format "HH:mm:ss.fff"
    $Date2 = Get-Date -Format "MM-dd-yyyy"

    if ($Append)
    {
        "<![LOG[$Message]LOG]!><time=`"$Date1+$($TimeZoneBias.bias)`" date=`"$Date2`" component=`"$LogFileComponent`" context=`"`" type=`"$Severity`" thread=`"-1`" file=`"$Program`">" | Out-File -FilePath $FilePath -Encoding utf8 -Force -Confirm:$False -Append
    }
    else
    {
        "<![LOG[$Message]LOG]!><time=`"$Date1+$($TimeZoneBias.bias)`" date=`"$Date2`" component=`"$LogFileComponent`" context=`"`" type=`"$Severity`" thread=`"-1`" file=`"$Program`">" | Out-File -FilePath $FilePath -Encoding utf8 -Force -Confirm:$False
    }
}

# This script needs to be started with administrative privileges
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
    Write-Error "This script must be started with administrative privileges!"
}

# Check for Hyper-V Management Tools
if (!(Get-Command "Get-VMHost" -ErrorAction SilentlyContinue))
{
    Write-Error "This script requires the following features to be installed: 'Hyper-V Management Tools' and 'Hyper-V Platform'"
}

# Check for Hyper-V Platform
try
{
    $VMHost = Get-VMHost -ErrorAction Stop
}
catch
{
    Write-Error "This script requires the following features to be installed: 'Hyper-V Management Tools' and 'Hyper-V Platform'"
}

# Make sure $LogFolder exists
if (!(Test-Path $LogFolder))
{
    New-Item -Path $LogFolder -ItemType Directory -Force
}

# Path to log file
[string]$LogFile = "$LogFolder\ExchangeLabEnvironment_$((Get-Date -Format "ddMMyy")).log"

# Write start section to log file
Write-Log -LogPath $LogFile -Message "########################################################################################" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "Script started $((Get-Date -Format "dd.MM.yyyy HH:mm:ss")) - By: $(([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).Identity.Name)" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "########################################################################################" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "Variables used:" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"

# Where is this hosts default Virtual Hard Drive path
[string]$VHDPath = (Get-VMHost).VirtualHardDiskPath

# Write variables section to log file
Write-Log -LogPath $LogFile -Message "Exchange server count: $NumberOfExchangeServers" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "Virtual hard drive type: $VirtualHardDriveType" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "Differencing Parent Path: $DifferencingParentPath" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "Source disk: $SourceDisk" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "Virtual hard drive path: $VHDPath" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "System drive size: $SizeSystemDrive" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "Data drive size: $SizeDataDrive" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "Memory size: $SizeRam" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "Virtual machine name: $VirtualMachineName" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "Virtual machine network interface name: $VirtualMachineNetworkInterfaceName" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
Write-Log -LogPath $LogFile -Message "########################################################################################" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"

########### Create hard drives ###########

if ($VHDType.ToLower() -eq "differencing")
{
    if (!$DifferencingParentPath)
    {
        Write-Log -LogPath $LogFile -Message "ERROR: DifferencingParentPath is needed when creating a Differencing VHD." -Severity 3 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
        Write-Error "DifferencingParentPath is needed when creating a Differencing VHD."
    }

    for ($i = 0; $i -lt $NumberOfExchangeServers; $i++)
    {
        # TODO: Add 0 to incremental number if less than 10
        try
        {
            New-VHD -ParentPath $DifferencingParentPath -Path ("$VHDPath\$VMName$(($i + 1)).vhdx") -Differencing -ErrorAction Stop
            Write-Log -LogPath $LogFile -Message "Differencing virtual hard drive '$("$VHDPath\$VMName$(($i + 1)).vhdx")' created" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
        }
        catch
        {
            Write-Log -LogPath $LogFile -Message "ERROR: Failed to create differencing virtual hard drive: $_" -Severity 3 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
            Write-Error "Failed to create differencing virtual hard drive: $_"
        }
    }
}

########### Create virtual machines ###########

for ($i = 0; $i -lt $NumberOfExchangeServers; $i++)
{
    # TODO: Add 0 to incremental number if less than 10
    try
    {
        New-VM -Name ("$VHDPath\$VMName$(($i + 1))") -MemoryStartupBytes $SizeRam -VHDPath ("$VHDPath\$VMName$(($i + 1)).vhdx")
        Write-Log -LogPath $LogFile -Message "Virtual machine '$("$VHDPath\$VMName$(($i + 1))")' created" -Severity 1 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
    }
    catch
    {
        Write-Log -LogPath $LogFile -Message "ERROR: Failed to create virtual machine: $_" -Severity 3 -Program "New-ExchangeLabEnvironment.ps1" -LogFileComponent "New-ExchangeLabEnvironment.ps1"
        Write-Error "Failed to create differencing virtual machine: $_"
    }
}