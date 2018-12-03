<# 
.SYNOPSIS
    Copy config files to excel

.DESCRIPTION
    Copy all files from specified directory to an excel file, fill headers with hostname
    Config files must be with extension .conf, .txt or .log
    
.NOTES 
    Author: Krzysztof Zaleski (cshyshtof@gmail.com)
    Date: 2017.12.13
    Version: 1.1
    Disclaimer: This script is provided as-is without any support or responsibility.
                You may copy and modify this script freely as long as you keep
                the reference to the original author.
    
.LINK 
    http://ccie24081.wordpress.com
#>

Clear-Host

#========================================================================================================
# Settable variables
#========================================================================================================

$WrkshtName = 'Configs'
$XlsFile = 'Configs.xlsx'
$XlsDir = '' # Current dir will be used if empty
$ConfigDir = 'c:\devops\collect-config\configs\'


#========================================================================================================
# Non-settable variables
#========================================================================================================

$CountFile = 1
$HeaderColumns = 0
$HeaderRows = 4
$HostnameRow = 1
$VersionRow = 2
$MgmtIpRow = 3
$RegexHostname = [regex]"^hostname (.+)"
$RegexStartOfConfig = [regex]"^version (.+)"
$FreezeRow = $HeaderRows + 1


#========================================================================================================
# Functions
#========================================================================================================

function Get-ScriptDirectory
{
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value;

    if($Invocation.PSScriptRoot)
    {
        $Invocation.PSScriptRoot;
    }
    Elseif($Invocation.MyCommand.Path)
    {
        Split-Path $Invocation.MyCommand.Path
    }
    else
    {
        $Invocation.InvocationName.Substring(0, $Invocation.InvocationName.LastIndexOf("\"));
    }
}


#========================================================================================================
# Check directories and files
#========================================================================================================

if (! $XlsDir)
{
    $XlsDir = Get-ScriptDirectory
}
$XlsPath = "$($XlsDir)\$($XlsFile)"

if (! (Test-Path $XlsDir -PathType Container))
{
    Write-Host "No such directory (XlsDir): $($XlsDir)"
    exit
}
else
{
    if (! (Test-Path $XlsPath -PathType Leaf))
    {
        Write-Host "File will be saved to: $($XlsPath)"
    }
    else
    {
        Write-Host "File already exists (XlsPath): $($XlsPath)"
        exit
    }
}

if (! (Test-Path $ConfigDir -PathType Container))
{
    Write-Host "No such directory (ConfigDir): $($ConfigDir)"
    exit
}
else
{
    $ConfigFiles = dir $ConfigDir\* -Include *.txt, *.log, *.conf
    $CountAllFiles = ($ConfigFiles | measure).Count
    Write-Host "Fetched configs in directory: $($ConfigDir) [$($CountAllFiles) files]"

    if ($CountAllFiles -eq 0)
    {
        Write-Host "Nothing to be done... exiting"
        exit
    }
}


#========================================================================================================
# Prepare excel spreadsheets
#========================================================================================================

$XlsStartProcs = Get-Process | Where {($_.Name.ToLower() -eq "excel")}
$InvExcel = New-Object -ComObject excel.application
$InvExcel.Visible = $True
$InvExcel.DisplayAlerts = $False
$InvExcel.UserControl = $False
$InvExcel.Interactive = $False
$InvWorkbook = $InvExcel.Workbooks.Add()
$InvWrksht = $InvWorkbook.Worksheets.Add()
$InvWrksht.Name = [string]$WrkshtName
$InvWrksht.Activate()
$InvWorkbook.Worksheets('Arkusz1').Delete()
$XlsEndProcs = Get-Process | Where {($_.Name.ToLower() -eq "excel")}


#========================================================================================================
# Inject files
#========================================================================================================

#$FirstCol = $HeaderColumns + 1
#$InvWrksht.Cells.Item(1, $FirstCol) = 'HOSTNAME'
#$InvWrksht.Cells.Item(2, $FirstCol) = 'MODEL'
#$InvWrksht.Cells.Item(3, $FirstCol) = 'IOS VER'
#$InvWrksht.Cells.Item(4, $FirstCol) = 'MGMT IP'

foreach ($ConfigFile in $ConfigFiles)
{
    Write-Progress -Id 1 -Activity "Injecting files" -Status "File $($CountFile) of $($CountAllFiles): $($ConfigFile)" -PercentComplete (($CountFile++ / $CountAllFiles) * 100);

    $Row = $HeaderRows + 1
    $InsideConfig = 0
    $RegexEndOfConfig = [regex]"---bogus---"

    foreach ($ConfigLine in Get-Content "$($ConfigFile)")
    {
        if ($ConfigLine -match $RegexStartOfConfig)
        {
            $InsideConfig = 1
            $Column = $HeaderColumns + $CountFile - 1
            $ConfigVersion = $matches[1]
            $InvWrksht.Cells.Item($VersionRow, $Column) = $ConfigVersion
        }

        if ($ConfigLine -match $RegexHostname)
        {
            $ConfigHostname = $matches[1]

            Write-Progress -Id 3 -ParentId 2 -Activity "Found hostname: " -Status $ConfigHostname

            $InvWrksht.Cells.Item($HostnameRow, $Column) = $ConfigHostname
            $RegexEndOfConfig = [regex]"^$($ConfigHostname)"
        }

        if ($ConfigLine -match $RegexEndOfConfig)
        {
            break
        }

        if ($InsideConfig)
        {
            $InvWrksht.Cells.Item($Row++, $Column) = $ConfigLine
        }        
    }
}


#========================================================================================================
# Add formating
#========================================================================================================

$MainRange = $InvWrksht.UsedRange.Cells
$MainRange.WrapText = $False
$MainRange.EntireColumn.ColumnWidth = 42
$MainRange.EntireRow.AutoFit() | Out-Null
#$MainRange.Borders.LineStyle = 1 # Solid
#$MainRange.Borders.Color = 0 # Black
#$MainRange.Borders.Weight = 2 # Thin
$MainRange.Font.Size = 9
$MainRange.Font.Name = "Courier New"

$HeadRange = $InvWrksht.Range("1:$($HeaderRows)")
$HeadRange.Interior.ColorIndex = 6 # Yellow
$HeadRange.Font.Bold = $True
$HeadRange.Font.ColorIndex = 0 # Black
$HeadRange.HorizontalAlignment = -4108 # Center

$InvExcel.Rows.Item("$($FreezeRow):$($FreezeRow)").Select() | Out-Null
$InvExcel.ActiveWindow.FreezePanes = $True

#========================================================================================================
# Excel save and cleanup
#========================================================================================================

$InvWorkbook.SaveAs($XlsPath)
$InvExcel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($InvExcel) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($InvWorkbook) | Out-Null
Remove-Variable InvExcel | Out-Null
Remove-Variable InvWorkbook | Out-Null

if ($XlsStartProcs -gt $Null)
{
    $CompareProcs = Compare-Object $XlsStartProcs $XlsEndProcs
    $XlsProcesses  = $CompareProcs | % {If($_.SideIndicator -match "=>") {$_.InputObject}}
}
else
{
     $XlsProcesses = $XlsEndProcs
}
 
Stop-Process -Id $XlsProcesses.Id | Out-Null