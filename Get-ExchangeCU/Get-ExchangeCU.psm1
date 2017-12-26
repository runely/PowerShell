<#
.Synopsis
   Download Cumulative Update for Exchange 2013/2016
.DESCRIPTION
   Download Cumulative Update for Exchange 2013/2016 to all or specified servers
.EXAMPLE
   Get-ExchangeCU [[-Version '2013_CUn'] [-Name Server1,Server2] [-Directory 'C$\Source'] [-RemoveTemp]]
   Download latest Cumulative Update for installed Exchange version (Exchange Management Shell required)
.EXAMPLE
   Get-MailboxServer | Get-ExchangeCU -Version '2013_CUn'
   Download Cumulative Update to given Exchange servers at default location(\\Server\C$\Source).
.EXAMPLE
   Get-ExchangeCU -Version '2013_CUn' -Name Server1,Server2
   Download Cumulative Update to given Exchange servers at default location(\\Server\C$\Source).
.EXAMPLE
   Get-MailboxServer | Get-ExchangeCU -Version '2013_CUn' [[-Directory 'C$\Source'] [-RemoveTemp]]
   Download Cumulative Update to given Exchange servers at specified location and remove temporary downloaded file
.EXAMPLE
   Get-ExchangeCU -Version '2013_CUn' -Name Server1,Server2 [[-Directory 'C$\Source'] [-RemoveTemp]]
   Download Cumulative Update to given Exchange servers at specified location and remove temporary downloaded file
.NOTES
   Author: Rune Moskvil Lyngås
   Version history:
   1.0.4
        - Added url for Exchange 2013 CU19 and Exchange 2016 CU8
        - Version parameter is no longer mandatory. If omitted, the latest CU for the installed Exchange version will be downloaded (requires Exchange Management Shell installed)
        - Name parameter is no longer mandatory. If omitted, CU will be downloaded to all Exchange servers (requires Exchange Management Shell installed)
   1.0.3
        - Added download urls into the module. No need to find them yourself anymore! Woho! :)
   1.0.2
        - Converted to a function. This makes it easier to import it through a powershell profile
   1.0.1
        - Fixed a bug where file would not be copied to all given servers when parameter $Name were given explicitly. This would not happen if names were pipelined in
        - Cleaned up the script a bit
        - More info are written to console
   1.0.0
        - Initial release
   
   Last updated: 21.12.2017
#>
Function Get-ExchangeCU
{
    [CmdletBinding(SupportsShouldProcess = $True)]
    Param(
        [Parameter(HelpMessage = "Exchange CU version to download")]
        [ValidateSet("2013_CU1", "2013_CU2", "2013_CU3", "2013_CU4/SP1", "2013_CU5", "2013_CU6", "2013_CU7", "2013_CU8", "2013_CU9", "2013_CU10", "2013_CU11", "2013_CU12", "2013_CU13", "2013_CU14", "2013_CU15", "2013_CU16", "2013_CU17", "2013_CU18", "2013_CU19", "2016_CU1", "2016_CU2", "2016_CU3", "2016_CU4", "2016_CU5", "2016_CU6", "2016_CU7", "2016_CU8")]
        [string]$Version,

        [Parameter(ValueFromPipelineByPropertyName = $True, HelpMessage = "Exchange servers. Given as an array: Server1,Server2")]
        [string[]]$Name,

        [Parameter(HelpMessage = "Download location on Exchange server. Given as an UNC path without the server name. Example: 'C$\Source' or 'Source'")]
        [ValidateNotNullOrEmpty()]
        [string]$Directory = "C$\Source",

        [Parameter(HelpMessage = "Give this switch to remove Cumulative Update from temp folder when script is done")]
        [switch]$RemoveTemp
    )

    begin
    {
        # if version is not set, choose the latest CU for the installed Exchange version
        if (!$Version)
        {
            # import Exchange Management Shell module and connect
            if (!(Get-Command "Get-ExchangeServer" -ErrorAction SilentlyContinue))
            {
                Write-Host "Exchange Management Shell is required to find installed Exchange version. Please start script in Exchange Management Shell or fill out Version parameter." -ForegroundColor Red
                break;
            }

            # get installed Exchange version
            $Servers = (Get-ExchangeServer | Select -ExpandProperty AdminDisplayVersion)

            # choose latest CU for installed Exchange version
            try
            {
                if ($Servers)
                {
                    if ($Servers.Count -gt 1)
                    {
                        if ($Servers[0].GetType().FullName -ne "Microsoft.Exchange.Data.ServerVersion")
                        {
                            Write-Host "Exchange Management Shell is required to find installed Exchange version. Please start script in Exchange Management Shell or fill out Version parameter." -ForegroundColor Red
                            break;
                        }

                        if ($Servers[0].Major -eq 15 -and $Servers[0].Minor -eq 0)
                        {
                            $Version = Get-ExchangeCUUri -Latest "2013"
                        }
                        elseif ($Servers[0].Major -eq 15 -and $Servers[0].Minor -eq 1)
                        {
                            $Version = Get-ExchangeCUUri -Latest "2016"
                        }
                    }
                    else
                    {
                        if ($Servers.GetType().FullName -ne "Microsoft.Exchange.Data.ServerVersion")
                        {
                            Write-Host "Exchange Management Shell is required to find installed Exchange version. Please start script in Exchange Management Shell or fill out Version parameter." -ForegroundColor Red
                            break;
                        }

                        if ($Servers.Major -eq 15 -and $Servers[0].Minor -eq 0)
                        {
                            $Version = Get-ExchangeCUUri -Latest "2013"
                        }
                        elseif ($Servers.Major -eq 15 -and $Servers[0].Minor -eq 1)
                        {
                            $Version = Get-ExchangeCUUri -Latest "2016"
                        }
                    }
                }
                else
                {
                    Write-Host "Command '(Get-ExchangeServer | Select -ExpandProperty AdminDisplayVersion)' did not return any results." -ForegroundColor Red
                    break;
                }
            }
            catch
            {
                Write-Host "Command '(Get-ExchangeServer | Select -ExpandProperty AdminDisplayVersion)' failed: ($_)" -ForegroundColor Red
                break;
            }
        }

        # get download link
        Write-Host "'$Version' - Getting download link : " -ForegroundColor Cyan -NoNewline
        [string]$Url = Get-ExchangeCUUri -Version $Version

        if ($Url -eq $null -or $Url -eq "" -or $Url.ToUpper() -eq "N/A")
        {
            Write-Host "Download link not available anymore. Please try a different version" -ForegroundColor Red
            break;
        }
        else
        {
            Write-Host "'$Url'" -ForegroundColor Green
        }

        # get filename of CU
        $FileName = Split-Path $Url -Leaf

        # create TempFile variable for CU
        $TempFile = $env:TEMP + "\$FileName"

        # download CU to $TempFile
        try
        {
            Write-Host "Downloading '$FileName' to '$TempFile' : " -ForegroundColor Cyan -NoNewline

            if (!(Test-Path $TempFile -ErrorAction Stop))
            {
                Start-BitsTransfer -Destination $TempFile -Source $Url -Description "Downloading $FileName to $TempFile" -ErrorAction Stop
                Write-Host "OK" -ForegroundColor Green
            }
            else
            {
                Write-Host "Already downloaded!" -ForegroundColor Green
            }
        }
        catch
        {
            Write-Host "Failed ($_)" -ForegroundColor Red
            break;
        }
    }

    process
    {
        # if not name is given, get all Exchange Servers
        if (!$Name)
        {
            # import Exchange Management Shell module and connect
            if (!(Get-Command "Get-ExchangeServer" -ErrorAction SilentlyContinue))
            {
                Write-Host "Exchange Management Shell is required to find installed Exchange servers. Please start script in Exchange Management Shell or fill out Version parameter." -ForegroundColor Red
                break;
            }

            # get installed Exchange servers
            $Name = (((Get-ExchangeServer | Select -ExpandProperty Name) -join ",") -split ",")
        }

        foreach ($Server in $Name)
        {
            $ServerPath = "\\$Server\$Directory"

            # make sure path exists on $Server
            try
            {
                Write-Host "Creating folder '$ServerPath' : " -ForegroundColor Cyan -NoNewline

                if (!(Test-Path $ServerPath -ErrorAction Stop))
                {
                    New-Item -Path $ServerPath -ItemType Directory | Out-Null
                    Write-Host "OK" -ForegroundColor Green
                }
                else
                {
                    Write-Host "Already exists." -ForegroundColor Green
                }
            }
            catch
            {
                Write-Host "Failed ($_)" -ForegroundColor Red
                continue;
            }
        
            # copy CU to server
            $ServerPath = "\\$Server\$Directory\$FileName"
            try
            {
                Write-Host "Copying '$FileName' to '$ServerPath' : " -ForegroundColor Cyan -NoNewline

                if (!(Test-Path $ServerPath -ErrorAction Stop))
                {
                    Copy-Item -Path $TempFile -Destination $ServerPath -Force
                    Write-Host "OK" -ForegroundColor Green
                }
                else
                {
                    Write-Host "Already exists." -ForegroundColor Green
                }
            }
            catch
            {
                Write-Host "Failed ($_)" -ForegroundColor Red

                if ($RemoveTemp)
                {
                    $PSBoundParameters.Remove('RemoveTemp') | Out-Null
                    Write-Host "RemoveTemp switch was removed due to an error with copying '$FileName' to '$ServerPath'" -ForegroundColor Yellow
                }
            }
        }
    }

    end
    {
        if ($RemoveTemp)
        {
            Write-Host "Removing temporary file '$TempFile' : " -ForegroundColor Cyan -NoNewline

            try
            {
                Remove-Item -Path $TempFile -Force -Confirm:$False
                Write-Host "OK" -ForegroundColor Green
            }
            catch
            {
                Write-Host "Failed ($_)" -ForegroundColor Red
            }
        }
    }
}

Function Get-ExchangeCUUri
{
    param(
        [Parameter(Mandatory = $True, ParameterSetName = "Version")]
        [string]$Version,

        [Parameter(Mandatory = $True, ParameterSetName = "Latest")]
        [ValidateSet("2013", "2016")]
        [string]$Latest
    )

    # Exchange version table
    $ExchangeTable = @()
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU1"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU2"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU3"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU4/SP1"; Value = "https://download.microsoft.com/download/8/4/9/8494E4ED-8FA8-40CA-9E89-B9317995AD7E/Exchange2013-x64-SP1.exe" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU5"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU6"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU7"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU8"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU9"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU10"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU11"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU12"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU13"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU14"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU15"; Value = "https://download.microsoft.com/download/3/A/5/3A5CE1A3-FEAA-4185-9A27-32EA90831867/Exchange2013-x64-cu15.exe" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU16"; Value = "https://download.microsoft.com/download/7/B/9/7B91E07E-21D6-407E-803B-85236C04D25D/Exchange2013-x64-cu16.exe" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU17"; Value = "https://download.microsoft.com/download/D/E/1/DE1C3D22-28A6-4A30-9811-0A0539385E51/Exchange2013-x64-cu17.exe" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU18"; Value = "https://download.microsoft.com/download/5/9/8/598B1735-BC2E-43FC-88DD-0CDFF838EE09/Exchange2013-x64-cu18.exe" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2013_CU19"; Value = "https://download.microsoft.com/download/3/A/4/3A4E9E23-E698-477D-B1E3-CA235CE3DB7C/Exchange2013-x64-cu19.exe" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2016_CU1"; Value = "https://download.microsoft.com/download/6/4/8/648EB83C-00F9-49B2-806D-E46033DA4AE6/ExchangeServer2016-CU1.iso" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2016_CU2"; Value = "https://download.microsoft.com/download/C/6/C/C6C10C1B-EFD8-4AE7-AEE1-C04F45869F5D/ExchangeServer2016-x64-CU2.iso" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2016_CU3"; Value = "N/A" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2016_CU4"; Value = "https://download.microsoft.com/download/B/9/F/B9F59CF4-7C60-49EF-8A5B-8C2B7991FA86/ExchangeServer2016-x64-cu4.iso" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2016_CU5"; Value = "https://download.microsoft.com/download/A/A/7/AA7F69B2-9E25-4073-8945-E4B16E827B7A/ExchangeServer2016-x64-cu5.iso" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2016_CU6"; Value = "https://download.microsoft.com/download/2/D/B/2DB1EEA2-CD9B-48F1-8235-1C9B82D19D68/ExchangeServer2016-x64-cu6.iso" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2016_CU7"; Value = "https://download.microsoft.com/download/0/7/4/074FADBD-4422-4BBC-8C04-B56428667E36/ExchangeServer2016-x64-cu7.iso" })
    $ExchangeTable += (New-Object PSObject -Property @{ Name = "2016_CU8"; Value = "https://download.microsoft.com/download/1/F/7/1F777B44-32CB-4F3D-B486-3D0F566D79A9/ExchangeServer2016-x64-cu8.iso" })

    if ($Version)
    {
        $Uri = ($ExchangeTable | Where { $_.Name -eq $Version } | Select -ExpandProperty Value)
    }
    elseif ($Latest)
    {
        $Uri = ($ExchangeTable | Where { $_.Name -like "$Latest*" } | Select -Last 1 | Select -ExpandProperty Name)
    }

    if ($Uri)
    {
        return $Uri
    }
    else
    {
        return "N/A"
    }
}