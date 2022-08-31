<# : file_splitter.bat
:: Select your TXT file and split in multiple files with max 30k lines

@echo off
setlocal EnableDelayedExpansion

for /f "delims=" %%I in ('powershell -noprofile "iex (${%~f0} | out-string)"') do (
    echo %%~I
)
goto :EOF


REM : end Batch portion / begin PowerShell hybrid chimera #>

Add-Type -AssemblyName System.Windows.Forms
$f = new-object Windows.Forms.OpenFileDialog
$f.InitialDirectory = pwd
$f.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
$f.ShowHelp = $true
$f.Multiselect = $false
[void]$f.ShowDialog()
if ($f.Multiselect) { $f.FileNames } else { $f.FileName }

$outDir = Split-Path -Path $f.FileName
$rootName = (Split-Path -Path $f.FileName -Leaf).Split(".")[0];
$ext = (Split-Path -Path $f.FileName -Leaf).Split(".")[1];
$nRows = 30000

$reader = new-object System.IO.StreamReader($f.FileName)
$count = 1
$rowCount = 0
$fileName = "{0} - part{1}.{2}" -f ($rootName, $count, $ext)
$header = $reader.ReadLine()

while(($line = $reader.ReadLine()) -ne $null)
{   
    if($rowCount -eq 0) {
        Add-Content -path $fileName -value $header
    }
    Add-Content -path $fileName -value $line
    ++$rowCount
    if($rowCount -ge $nRows)
    {
        ++$count
        $rowCount = 0
        $fileName = "{0} - part{1}.{2}" -f ($rootName, $count, $ext)
    }
}

$reader.Close()