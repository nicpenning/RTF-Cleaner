#Created by Nicholas Penning
#Last modified: 11/03/2017

#Use this tool to statically find URLs in obfuscated RTF documents. 
#Usage: .\rtfCleanerUniversal.ps1 path-to-file
#Example: .\rtfCleanerUniversal.ps1 'C:\Users\yams\Downloads\invoiceForYams.doc'
#If the built in regexers don't find anything just enter what you find in the rtf file, IE pntxtb, lchars, etc... 
#Looking to build in functionality to find these and do it automatically. Its manual for now though.
#This works for any file type that contains {\parameters} to clean.
#Use at your own risk!

Param($filePath)
Write-Host "Analyzing:"$filePath
$fileTxt = [IO.File]::ReadAllText($filePath)

$regExPNTXTa = '\{\\pntxta\s*([^}]*?)\s*}'
$regExLChars = '\{\\lchars\s*([^}]*?)\s*}'
$regExForHex = '(010000020900000001[A-Fa-f0-9]{280})'
$RegExLink = '(http|ftp|https):\/\/([\w_-]+(?:(?:\.[\w_-]+)+))([\w.,@?^=%&:/~+#-]*[\w@?^=%&/~+#-]?)'

#Uncomment lines 13-19 and 73 to prompt a dialog box for the file instead of using this tool as command line. (Only works on Windows)
<#[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $initialDirectory
$OpenFileDialog.filter = "DOC (*.doc)| *.doc|RTF (*.rtf)|*.rtf|TXT (*.txt)|*.txt"
$OpenFileDialog.ShowDialog() | Out-Null
$OpenFileDialog.FileName
#>

#Find Matches to {\* txt that is used for hiding hex}
function findMatches{
    $matchPNTXTa = $fileTxt | Select-String -Pattern $regExPNTXTa -AllMatches
    $matchLChars = $fileTxt | Select-String -Pattern $regExLChars -AllMatches
    

    if($matchPNTXTa.Matches.Count -ge 1){
        Write-Host $matchPNTXTa.Matches.Count of $regExPNTXTa Found!
        filterText($regExPNTXTa)
    }elseif($matchLChars.Matches.Count -ge 1){
        Write-Host $matchLChars.Matches.Count of $regExLChars Found!
        filterText($regExLChars)
    }else{
        #If no regular expressions trigger, ask for basic search
        Write-Host 'Enter your custom parameter to filter on, for example: lchars or pntxta' -ForegroundColor Yellow
        $customParameter = Read-Host
        $regExCustom = '\{\\'+$customParameter+'\s*([^}]*?)\s*}'
        $matchCustom = $fileTxt | Select-String -Pattern $regExCustom -AllMatches
        Write-Host $matchCustom.Matches.Count of $regExCustom Found!
        filterText($regExCustom)
    }
    
}

function filterText($regEx){
    $filter = ($fileTxt -replace $regEx, "")
    $whiteSpaceFiltered = ($filter -replace "\s", "")
    $whiteSpaceFiltered2 = ($whiteSpaceFiltered -replace "\n", "")
    $cleaned = ($whiteSpaceFiltered2 -replace "\r", "")

    findURL($cleaned)
}
#Decoded and Find URL
function findURL($textToDecode){
    $decodeMe = $textToDecode | Select-String -Pattern $regExForHex

    #Uncomment below to see if the hex is being found
    #$decodeMe.Matches.Value

    #Decode Hex
    $found = -join ($decodeMe.Matches.Value -split '(..)' | ? { $_ } | % { [char][convert]::ToUInt32($_,16) })

    #Remove NULL characters
    $finalCleaning = $found -replace "`0",''

    #Extract URL
    $foundURL = $finalCleaning | Select-String -Pattern $RegExLink -AllMatches
    if($foundURL.Matches.Value -ge 1){
        Write-Host 'URL Found!' -fore green
        $foundURL.Matches.Value
    }else{Write-Host 'Sorry, URL could not be extracted' -fore red}
}

#$fileTxt = [IO.File]::ReadAllText($OpenFileDialog.FileName)
findMatches
