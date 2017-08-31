<#
.NOTES
	Name: PF-Find-Clean.ps1
	Original Author: Chris Heilman
	Requires: Exchange Management Shell (Exchange Server 2010) and administrator rights on the Exchange server and Public Folders.
	Version: 1.0 -- 08/31/2017

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
	
	
.SYNOPSIS
	Checks the target Exchange server for illegal characters and leading or trailing spaces in Mail-enabled Public Folder aliases.

.DESCRIPTION
	Exchange 2000 and 2003 would end up allowing illegal characters (as of Exchange 2007 and up) in the Alias attribute. When migrating legacy public folders from 2007/2010 to Exchange 2013/2016 we encounter these issues.
	
.PARAMETER Repair
	This required parameter is a $True / $False switch. When set to $False this will just run a scan of the public folders and output what it finds. When set to $True this will scan and then fix the public folders that have illegal characters or leading and trailing spaces.
	
.EXAMPLE
	.\PF-Clean-Fix.ps1 -Repair $false
	Run a scan against the mail-enabled public folders
	
.EXAMPLE
	.\PF-Clean-Fix.ps1 -Repair $true
	Run a scan against the mail-enabled public folders and proceed to fix them by changing the offending characters to a hyphen "-"
#>
#------------------------------------------
# Setting Parameter

Param(
  [Parameter(Mandatory=$True, Position=1)]
   [string]$Repair
   )

#------------------------------------------
# Setting our main variable 

$MPF = Get-MailPublicFolder -resultSize Unlimited | where {$_.Alias.ToCharArray() -contains ' ' -or $_.Alias.ToCharArray() -contains '@' -or $_.Alias.ToCharArray() -contains ',' -or $_.Alias.ToCharArray() -contains ':' -or $_.Alias.ToCharArray() -contains ';' -or $_.Alias.ToCharArray() -contains '(' -or $_.Alias.ToCharArray() -contains ')' -or $_.Alias.ToCharArray() -contains '\'}

#------------------------------------------
#newLine shortcut
$script:nl = "`r`n"
$nl

Clear-Host


#------------------------------------------
# Starting Transcript
function transcriptStart

{
Start-Transcript PublicFolder-Find-Clean.txt -append
}


#------------------------------------------
# Stopping Transcript
function transcriptStop

{
Stop-Transcript
}


#------------------------------------------
function beforeOutput
{
Get-PublicFolder \ -recurse -ResultSize Unlimited | where{$_.MailEnabled -eq "True"} | Get-MailPublicFolder -resultSize Unlimited | Out-File Mail-Public-Folders-Before.csv
}


#------------------------------------------
function afterOutput
{
Get-PublicFolder \ -recurse -ResultSize Unlimited | where{$_.MailEnabled -eq "True"} | Get-MailPublicFolder -resultSize Unlimited | Out-File Mail-Public-Folders-After.csv
}


#------------------------------------------
# Grabbing and Sorting through the Mail-enabled Public Folders

function sortMailPF
{
$nl

$MPFCount = ($MPF).count

If ($MPFCount -eq $null){$MPFCount = 0}
Else {$MPFCount = $MPFCount}


Write-Host "----------------------------" -foregroundcolor Green

$nl
Write-Host "Found $MPFCount Mail-enabled Public Folders with Spaces or Bad Characters." -foregroundcolor Yellow
$nl

}

#------------------------------------------
# Grabbing and Sorting again for recheck

function sortMailPFagain
{

$MPFCount = ($MPF2).count

If ($MPFCount -eq $null){$MPFCount = 0}

Write-Host "----------------------------" -foregroundcolor Green

$nl
Write-Host "Found $MPFCount Mail-enabled Public Folders with Spaces or Bad Characters." -foregroundcolor Yellow
$nl

}

#------------------------------------------
# Showing what objects we have found:

function findItems
{

Write-Host "----------------------------" -foregroundcolor Green
$nl
Write-Host "What we found:" -foregroundcolor white
$nl

$space = $MPF | where {$_.alias -like '* *'}
$spaceCount = ($space).count
If ($spaceCount -eq $null){$spaceCount = 0}
Write-Host "Found $spaceCount Spaces in an Alias." -foregroundcolor Yellow
$nl

$comma = $MPF | where {$_.alias -like '*,*'}
$commaCount = ($comma).count
If ($commaCount -eq $null){$commaCount = 0}
Write-Host "Found $commaCount Commas in an Alias." -foregroundcolor Yellow
$nl

$at = $MPF | where {$_.alias -like '*@*'}
$atCount = ($at).count
If ($atCount -eq $null){$atCount = 0}
Write-Host "Found $atCount '@' in an Alias." -foregroundcolor Yellow
$nl

$open = $MPF | where {$_.alias -like '*(*'}
$openCount = ($open).count
If ($openCount -eq $null){$openCount = 0}
Write-Host "Found $openCount '(' in an Alias." -foregroundcolor Yellow
$nl

$close = $MPF | where {$_.alias -like '*)*'}
$closeCount = ($close).count
If ($closeCount -eq $null){$closeCount = 0}
Write-Host "Found $closeCount ')' in an Alias." -foregroundcolor Yellow
$nl

$colon = $MPF | where {$_.alias -like '*:*'}
$colonCount = ($colon).count
If ($colonCount -eq $null){$colonCount = 0}
Write-Host "Found $colonCount ':' in an Alias." -foregroundcolor Yellow
$nl

$semicolon = $MPF | where {$_.alias -like '*;*'}
$semicolonCount = ($semicolon).count
If ($semicolonCount -eq $null){$semicolonCount = 0}
Write-Host "Found $semicolonCount ';' in an Alias." -foregroundcolor Yellow
$nl

$backslash = $MPF | where {$_.alias -like '*\*'}
$backslashCount = ($backslash).count
If ($backslashCount -eq $null){$backslashCount = 0}
Write-Host "Found $backslashCount '\' in an Alias." -foregroundcolor Yellow
$nl

Write-Host "----------------------------" -foregroundcolor Green

}


#------------------------------------------
# Previewing our replacement of Special Characters with an Hypens "-"

function previewReplace 
{

foreach($pf in $MPF){
   $newAlias = $pf.alias
if($newAlias -ne $null){
        $newAlias = $newAlias.Trim()
	    $newAlias = $newAlias.Replace(' ', '-')
		$newAlias = $newAlias.Replace(',', '-')
	    $newAlias = $newAlias.Replace('@', '-')
	    $newAlias = $newAlias.Replace('(', '-')
	    $newAlias = $newAlias.Replace(')', '-')
	    $newAlias = $newAlias.Replace(':', '-')
	    $newAlias = $newAlias.Replace(';', '-')
	    $newAlias = $newAlias.Replace('\', '-')
		$newAlias = $newAlias.Trim()
			Write-Host("New Alias is now: {0}" -f $newAlias) -foregroundcolor Cyan 
}
else
{
	Write-host("Public Folder Aliases are empty") -foregroundcolor Green
	$nl
	Write-Host "----------------------------" -foregroundcolor Green
	$nl
        } 
    }
}


#------------------------------------------
# Replacement of Special Characters with a Hypens "-"

function replaceCharacters
{

foreach($pf in $MPF){
    $newAlias = $pf.alias
	if($newAlias -ne $null){
	    $newAlias = $newAlias.Trim()
	    $newAlias = $newAlias.Replace(' ', '-')
		$newAlias = $newAlias.Replace(',', '-')
	    $newAlias = $newAlias.Replace('@', '-')
	    $newAlias = $newAlias.Replace('(', '-')
	    $newAlias = $newAlias.Replace(')', '-')
	    $newAlias = $newAlias.Replace(':', '-')
	    $newAlias = $newAlias.Replace(';', '-')
	    $newAlias = $newAlias.Replace('\', '-')
		$newAlias = $newAlias.Trim()
			Write-Host("New Alias is now: {0}" -f $newAlias) -foregroundcolor Cyan 
			
		Set-MailpublicFolder -Identity $pf.identity -Alias $newAlias
		Start-Sleep -s 1 
}
else{
	Write-host("Public Folder Aliases are empty") -foregroundcolor Green
	$nl
	Write-Host "----------------------------" -foregroundcolor Green
	$nl
    }
}


$nl

Write-Host "----------------------------" -foregroundcolor Green
$nl

}


#------------------------------------------
# Replacement of Special Characters with an Hypens "-"

function finalOutput

{
Write-Host "Here's your Mail-enabled public folders now:" -ForegroundColor Yellow

Get-PublicFolder \ -recurse -ResultSize Unlimited | where{$_.MailEnabled -eq "True"} | Get-MailPublicFolder -resultSize Unlimited

$nl
}


#------------------------------------------
#--------------------------------------
# Body of Script
#--------------------------------------

function Main {

If ($Repair -eq $false)
	{
	transcriptStart
	$nl
    If($MPF -ne $null)
	{
		Write-Host "=====================================================================================" -foregroundcolor Green $nl
		Write-Host "Sorting through the Mail-enabled Public Folders to find Spaces and Special Characters" -foregroundcolor White $nl
		Write-Host "=====================================================================================" -foregroundcolor Green 

	sortMailPF
	findItems
	$nl
		write-host "Preview of the changes:" -foregroundcolor White
	$nl
	previewReplace
	$nl
		write-host "Please re-run script with -Repair $true to fix the found items" -foregroundcolor White
		$nl}
    Else {
	Write-Host "No Bad objects or spaces found!" $nl$nl "Ending Script" -foregroundcolor Green}
    $nl
	transcriptStop
	}
	
	
	
Elseif ($Repair -eq $true)
	{
	transcriptStart
	$nl
	If($MPF -ne $null)
	{
	$nl
		Write-host "Outputting Mail-enabled Public Folders before changes" -foregroundcolor DarkYellow
	beforeOutput
	$nl	
		Write-Host "============================" -foregroundcolor Green $nl
		Write-Host "Sorting through the Mail-enabled Public Folders to find Spaces and Special Characters" -foregroundcolor White $nl
		Write-Host "============================" -foregroundcolor Green 

	sortMailPF
	findItems
	$nl
		write-host "Preview of the changes:" -foregroundcolor White
	$nl
	replaceCharacters
		Write-Host "Checking Mail Public Folders again" -foregroundcolor White
    $MPF2 = Get-MailPublicFolder -resultSize Unlimited | where {$_.Alias.ToCharArray() -contains ' ' -or $_.Alias.ToCharArray() -contains '@' -or $_.Alias.ToCharArray() -contains ',' -or $_.Alias.ToCharArray() -contains ':' -or $_.Alias.ToCharArray() -contains ';' -or $_.Alias.ToCharArray() -contains '(' -or $_.Alias.ToCharArray() -contains ')' -or $_.Alias.ToCharArray() -contains '\'}
	sortMailPFagain
		Write-Host "Outputting Mail-enabled Public Folders after changes" -foregroundcolor DarkYellow
	$nl}
	Else {
	Write-Host "No Bad objects or spaces found!" $nl$nl "Ending Script" -foregroundcolor Green}
	$nl
	transcriptStop
	}
}

Main

#--------------------------------------
