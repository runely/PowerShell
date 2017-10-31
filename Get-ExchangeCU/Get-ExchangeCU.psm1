<#
.Synopsis
   Download Cumulative Update for Exchange 2013/2016 to specified servers
.DESCRIPTION
   Download Cumulative Update for Exchange 2013/2016 to specified servers
.EXAMPLE
   Get-MailboxServer | Get-ExchangeCU -Version '2013_CU18'
   Download Cumulative Update to given Exchange servers at default location(\\Server\C$\Source).
.EXAMPLE
   Get-ExchangeCU -Version '2013_CU18' -Name Server1,Server2
   Download Cumulative Update to given Exchange servers at default location(\\Server\C$\Source).
.EXAMPLE
   Get-MailboxServer | Get-ExchangeCU -Version '2013_CU18' [[-Directory 'C$\Source'] [-RemoveTemp]]
   Download Cumulative Update to given Exchange servers at specified location and remove temporary downloaded file
.EXAMPLE
   Get-ExchangeCU -Version '2013_CU18' -Name Server1,Server2 [[-Directory 'C$\Source'] [-RemoveTemp]]
   Download Cumulative Update to given Exchange servers at specified location and remove temporary downloaded file
.NOTES
   Author: Rune Moskvil Lyngås
   Version history:
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
   
   Last updated: 31.10.2017
#>
Function Get-ExchangeCU
{
    [CmdletBinding(SupportsShouldProcess = $True)]
    Param(
        [Parameter(Mandatory = $True, HelpMessage = "Exchange CU version to download")]
        [ValidateSet("2013_CU1", "2013_CU2", "2013_CU3", "2013_CU4/SP1", "2013_CU5", "2013_CU6", "2013_CU7", "2013_CU8", "2013_CU9", "2013_CU10", "2013_CU11", "2013_CU12", "2013_CU13", "2013_CU14", "2013_CU15", "2013_CU16", "2013_CU17", "2013_CU18", "2016_CU1", "2016_CU2", "2016_CU3", "2016_CU4", "2016_CU5", "2016_CU6", "2016_CU7")]
        [string]$Version,

        [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True, HelpMessage = "Exchange servers. Given as an array: Server1,Server2")]
        [string[]]$Name,

        [Parameter(HelpMessage = "Download location on Exchange server. Given as an UNC path without the server name. Example: 'C$\Source' or 'Source'")]
        [ValidateNotNullOrEmpty()]
        [string]$Directory = "C$\Source",

        [Parameter(HelpMessage = "Give this switch to remove Cumulative Update from temp folder when script is done")]
        [switch]$RemoveTemp
    )

    begin
    {
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
        if (!(Test-Path $TempFile))
        {
            Write-Host "Downloading '$FileName' to '$TempFile' : " -ForegroundColor Cyan -NoNewline
            try
            {
                Start-BitsTransfer -Destination $TempFile -Source $Url -Description "Downloading $FileName to $TempFile" -ErrorAction Continue
                Write-Host "OK" -ForegroundColor Green
            }
            catch
            {
                Write-Host "Failed ($_)" -ForegroundColor Red
                Return;
            }
        }
    }

    process
    {
        foreach ($Server in $Name)
        {
            $ServerPath = "\\$Server\$Directory"

            # make sure path exists on $Server
            if (!(Test-Path $ServerPath))
            {
                Write-Host "Creating folder '$ServerPath' : " -ForegroundColor Cyan -NoNewline

                try
                {
                    New-Item -Path $ServerPath -ItemType Directory | Out-Null
                    Write-Host "OK" -ForegroundColor Green
                }
                catch
                {
                    Write-Host "Failed ($_)" -ForegroundColor Red
                    Return;
                }
            }
        
            # copy CU to server
            $ServerPath = "\\$Server\$Directory\$FileName"
            if (!(Test-Path $ServerPath))
            {
                Write-Host "Copying '$FileName' to '$ServerPath' : " -ForegroundColor Cyan -NoNewline
            
                try
                {
                    Copy-Item -Path $TempFile -Destination $ServerPath -Force
                    Write-Host "OK" -ForegroundColor Green
                }
                catch
                {
                    Write-Host "Failed ($_)" -ForegroundColor Red

                    if ($RemoveTemp)
                    {
                        $PSBoundParameters.Remove('RemoveTemp') | Out-Null
                        Write-Verbose "RemoveTemp switch was removed due to an error with copying '$FileName' to '$ServerPath'"
                    }
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
        [Parameter(Mandatory = $True)]
        [string]$Version
    )

    # Exchange hash table
    $UriHashExchange = @{
        "2013_CU1"      = "N/A";
        "2013_CU2"      = "N/A";
        "2013_CU3"      = "N/A";
        "2013_CU4/SP1"  = "https://download.microsoft.com/download/8/4/9/8494E4ED-8FA8-40CA-9E89-B9317995AD7E/Exchange2013-x64-SP1.exe";
        "2013_CU5"      = "N/A";
        "2013_CU6"      = "N/A";
        "2013_CU7"      = "N/A";
        "2013_CU8"      = "N/A";
        "2013_CU9"      = "N/A";
        "2013_CU10"     = "N/A";
        "2013_CU11"     = "N/A";
        "2013_CU12"     = "N/A";
        "2013_CU13"     = "N/A";
        "2013_CU14"     = "N/A";
        "2013_CU15"     = "https://download.microsoft.com/download/3/A/5/3A5CE1A3-FEAA-4185-9A27-32EA90831867/Exchange2013-x64-cu15.exe";
        "2013_CU16"     = "https://download.microsoft.com/download/7/B/9/7B91E07E-21D6-407E-803B-85236C04D25D/Exchange2013-x64-cu16.exe";
        "2013_CU17"     = "https://download.microsoft.com/download/D/E/1/DE1C3D22-28A6-4A30-9811-0A0539385E51/Exchange2013-x64-cu17.exe";
        "2013_CU18"     = "https://download.microsoft.com/download/5/9/8/598B1735-BC2E-43FC-88DD-0CDFF838EE09/Exchange2013-x64-cu18.exe";
        "2016_CU1"      = "https://download.microsoft.com/download/6/4/8/648EB83C-00F9-49B2-806D-E46033DA4AE6/ExchangeServer2016-CU1.iso";
        "2016_CU2"      = "https://download.microsoft.com/download/C/6/C/C6C10C1B-EFD8-4AE7-AEE1-C04F45869F5D/ExchangeServer2016-x64-CU2.iso";
        "2016_CU3"      = "N/A";
        "2016_CU4"      = "https://download.microsoft.com/download/B/9/F/B9F59CF4-7C60-49EF-8A5B-8C2B7991FA86/ExchangeServer2016-x64-cu4.iso";
        "2016_CU5"      = "https://download.microsoft.com/download/A/A/7/AA7F69B2-9E25-4073-8945-E4B16E827B7A/ExchangeServer2016-x64-cu5.iso";
        "2016_CU6"      = "https://download.microsoft.com/download/2/D/B/2DB1EEA2-CD9B-48F1-8235-1C9B82D19D68/ExchangeServer2016-x64-cu6.iso";
        "2016_CU7"      = "https://download.microsoft.com/download/0/7/4/074FADBD-4422-4BBC-8C04-B56428667E36/ExchangeServer2016-x64-cu7.iso";
    }

    return $UriHashExchange.Item($Version)
}