<#
.Synopsis
   Download Cumulative Update for Exchange 2013/2016 by given url to specified servers
.DESCRIPTION
   Download Cumulative Update for Exchange 2013/2016 by given url to specified servers
.EXAMPLE
   Get-MailboxServer | ./Download-CU.ps1 -Url 'https://download.microsoft.com/download/5/9/8/598B1735-BC2E-43FC-88DD-0CDFF838EE09/Exchange2013-x64-cu18.exe'
   Download Cumulative Update to given Exchange servers at default location(\\Server\C$\Source).
.EXAMPLE
   ./Download-CU.ps1 -Url 'https://download.microsoft.com/download/5/9/8/598B1735-BC2E-43FC-88DD-0CDFF838EE09/Exchange2013-x64-cu18.exe' -Name Server1,Server2
   Download Cumulative Update to given Exchange servers at default location(\\Server\C$\Source).
.EXAMPLE
   Get-MailboxServer | ./Download-CU.ps1 -Url 'https://download.microsoft.com/download/5/9/8/598B1735-BC2E-43FC-88DD-0CDFF838EE09/Exchange2013-x64-cu18.exe' [[-Directory 'C$\Source'] [-RemoveTemp]]
   Download Cumulative Update to given Exchange servers at specified location and remove temporary downloaded file
.EXAMPLE
   ./Download-CU.ps1 -Url 'https://download.microsoft.com/download/5/9/8/598B1735-BC2E-43FC-88DD-0CDFF838EE09/Exchange2013-x64-cu18.exe' -Name Server1,Server2 [[-Directory 'C$\Source'] [-RemoveTemp]]
   Download Cumulative Update to given Exchange servers at specified location and remove temporary downloaded file
.NOTES
   Author: Rune Moskvil Lyngås
   Version history:
   1.0.1
        - Fixed a bug where file would not be copied to all given servers when parameter $Name were given explicitly. This would not happen if names were pipelined in
        - Cleaned up the script a bit
        - More info are written to console
   1.0.0
        - Initial release
   Last updated: 24.10.2017
#>
[CmdletBinding(SupportsShouldProcess = $True)]
Param(
    [Parameter(Mandatory = $True, HelpMessage = "Cumulative Update Url. Example: https://download.microsoft.com/download/5/9/8/598B1735-BC2E-43FC-88DD-0CDFF838EE09/Exchange2013-x64-cu18.exe")]
    [ValidatePattern("[-a-zA-Z0-9@:%_\+.~#?&//=]{2,256}\.[a-z]{2,4}\b(\/[-a-zA-Z0-9@:%_\+.~#?&//=]*)?")]
    [string]$Url,

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
